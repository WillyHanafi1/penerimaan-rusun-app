"""Recover explicit housing identities without guessing a tenant from an amount."""

import re

import pandas as pd

from kasda_matching import MONTH_ALIASES, MONTHS, VA_REFERENCE_PATTERN, _money, _text, _unit_code


def valid_housing_code(value):
    code = _text(value)
    return bool(re.fullmatch(r'0[123]0[1-4]0[1-5](?:0[1-9]|1\d|2[0-4])', code))


def parse_bank_identity(narration):
    text = re.sub(r'\s+', ' ', _text(narration).upper()).strip()
    compact = re.sub(r'\s+', '', text)
    result = {}
    if 'PAYMENTQRIS' in compact:
        return result  # Merchant/MPAN digits do not identify the tenant's VA.
    references = {m.group(1) for source in (text, compact)
                  for m in VA_REFERENCE_PATTERN.finditer(source) if valid_housing_code(m.group(2))}
    if len(references) == 1:
        reference = references.pop()
        return {'Kode_8_Digit': reference[4:12], 'Bulan': MONTHS[int(reference[12:14]) - 1],
                'Tahun': '20' + reference[14:16], 'Metode_Bayar': 'VA',
                'Sumber_Deteksi': 'Nomor VA pada rekening koran'}
    if len(references) > 1:
        return {'_identity_issue': 'Beberapa nomor VA berbeda dalam satu transaksi; perlu alokasi manual'}

    coordinates = re.search(
        r'\bHUNIAN\s+([A-D])\s*(IV|III|II|I|V|[1-5])\s*(\d{1,2})\s+'
        r'(?:RSN|RUSUN(?:AWA)?)\s+(CIGUGUR(?: TENGAH)?|CIBEUREUM|MELONG|LEUWIGAJAH)\b', text)
    codes = {m.group(0) for m in re.finditer(r'(?<!\d)0[123]\d{6}(?!\d)', compact)
             if valid_housing_code(m.group(0))} if 'SETORTUNAI' in compact else set()
    if coordinates:
        building, floor, unit, site = coordinates.groups()
        coordinate_code = _unit_code(site, building, floor, unit)
        if not valid_housing_code(coordinate_code):
            return {'_identity_issue': 'Nomor hunian pada rekening koran tidak valid; perlu cek manual'}
        codes.add(coordinate_code)
    if len(codes) > 1:
        return {'_identity_issue': 'Beberapa identitas hunian berbeda dalam satu transaksi; perlu cek manual'}
    if not codes:
        return result
    result['Kode_8_Digit'] = codes.pop()
    if 'SETORTUNAI' in compact:
        result['Metode_Bayar'] = 'TUNAI'
        result['Sumber_Deteksi'] = 'Kode hunian pada setoran tunai'
    if coordinates:
        result['Sumber_Deteksi'] = 'Hunian dan periode pada rekening koran'
        if 'REK JAMINAN' in text:
            result['Metode_Bayar'] = 'PEMINDAHBUKUAN'
        name = re.search(r'\bAN\s+(.+?)\s+HUNIAN\b', text)
        if name:
            result['_bank_name'] = name.group(1).title()
        period = re.search(r'\bSEWA\s+(.+?)(?:\s+AN\s+|\s+HUNIAN\b|$)', text)
        if period:
            aliases = '|'.join(sorted(MONTH_ALIASES, key=len, reverse=True))
            months = list(dict.fromkeys(MONTHS[MONTH_ALIASES[m.group(1).lower()] - 1]
                          for m in re.finditer(rf'(?<![A-Z])({aliases})(?![A-Z])', period.group(1), re.I)))
            years = set(re.findall(r'(?<!\d)20\d{2}(?!\d)', period.group(1)))
            if months and len(years) == 1:
                result.update(Bulan=', '.join(months), Tahun=years.pop())
                if len(months) > 1:
                    result['_identity_issue'] = (
                        f'Pembayaran beberapa periode ({result["Bulan"]} {result["Tahun"]}); '
                        'hunian terdeteksi, nominal per periode perlu dialokasikan manual')
    return result


def add_housing_columns(frame):
    frame = frame.copy()
    for column in ('Rusunawa', 'Gedung', 'Lantai', 'No Hunian'):
        if column not in frame:
            frame[column] = pd.Series('', index=frame.index, dtype=object)
    for index, row in frame.iterrows():
        code = _text(row.get('Kode_8_Digit'))
        if valid_housing_code(code):
            frame.loc[index, ['Rusunawa', 'Gedung', 'Lantai', 'No Hunian']] = [
                {'01': 'Cigugur Tengah', '02': 'Cibeureum', '03': 'Leuwigajah'}[code[:2]],
                chr(64 + int(code[2:4])), ('I', 'II', 'III', 'IV', 'V')[int(code[4:6]) - 1],
                int(code[6:8]),
            ]
    return frame


def enrich_bank_identities(detected, manual):
    detected_ids = set(detected['ID_Transaksi'])
    bank = pd.concat([detected, manual], ignore_index=True).sort_values('Urutan_Rekening_Koran', kind='stable')
    for column in ('Metode_Bayar', '_identity_issue', '_bank_name', 'Keterangan'):
        if column not in bank:
            bank[column] = pd.Series('', index=bank.index, dtype=object)
    for index, row in bank.iterrows():
        credit, debit = _money(row.get('Credit Transaction')), _money(row.get('Debit Transaction'))
        if credit is None or credit <= 0 or (debit is not None and debit > 0):
            detected_ids.discard(row['ID_Transaksi'])
            continue
        identity = parse_bank_identity(row.get('Narasi'))
        if identity.get('Kode_8_Digit'):
            # Explicit VA/coordinates are stronger than a digit substring in a narrative.
            for column, value in identity.items():
                bank.at[index, column] = value
            detected_ids.add(row['ID_Transaksi'])
        elif identity.get('_identity_issue'):
            bank.at[index, '_identity_issue'] = identity['_identity_issue']
            bank.at[index, 'Keterangan'] = identity['_identity_issue']
            bank.at[index, 'Kode_8_Digit'] = ''
            detected_ids.discard(row['ID_Transaksi'])
        elif row['ID_Transaksi'] in detected_ids and not identity.get('Kode_8_Digit'):
            # The legacy regex can pick eight digits out of a longer account/reference.
            bank.at[index, 'Kode_8_Digit'] = ''
            bank.at[index, 'Keterangan'] = 'Kode hunian tidak berdiri sendiri pada narasi; perlu cek manual'
            detected_ids.discard(row['ID_Transaksi'])
    bank = add_housing_columns(bank)
    selected = bank['ID_Transaksi'].isin(detected_ids)
    return bank[selected].reset_index(drop=True), bank[~selected].reset_index(drop=True)


def describe_manual_housing(frame):
    """Show housing coordinates while keeping unfilled master names blank."""
    frame = add_housing_columns(frame)
    if 'Nama Penghuni' not in frame:
        frame['Nama Penghuni'] = pd.Series('', index=frame.index, dtype=object)
    return frame
