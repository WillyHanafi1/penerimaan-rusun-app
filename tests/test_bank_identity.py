import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

import app
from bank_identity import parse_bank_identity


def transaction(narration, credit=400000, debit=''):
    return {'Narasi': narration, 'Posting Date': '08 Oct 2026',
            'Credit Transaction': credit, 'Debit Transaction': debit}


class BankIdentityTests(unittest.TestCase):
    def test_va_identifies_unit_and_rental_period_without_kasda(self):
        for narration in ('MBL-M001-246400 PAY VA MBL 1295030205241026/Nopi',
                          'ATM-JALN-303058 TRF DARI BCA Erni NO REK 1295020104170726'):
            with self.subTest(narration=narration):
                detected, manual, audit = app.process_bank_sources(pd.DataFrame([transaction(narration)]))
                self.assertEqual(len(detected), 1)
                self.assertTrue(manual.empty)
                self.assertIsNone(audit)
                expected = ('03020524', 'Oktober', 'Leuwigajah', 'B', 'V', 24) if 'Nopi' in narration else (
                    '02010417', 'Juli', 'Cibeureum', 'A', 'IV', 17)
                self.assertEqual(tuple(detected.iloc[0][['Kode_8_Digit', 'Bulan', 'Rusunawa', 'Gedung', 'Lantai', 'No Hunian']]), expected)
                self.assertEqual(detected.iloc[0]['Tahun'], '2026')

    def test_explicit_single_period_transfer_preserves_name_and_location(self):
        narration = 'H126/900/ND-14 PBY SEWA SEPTEMBER 2026 AN PENGHUNI CONTOH HUNIAN A I 12 RSN CIGUGUR TENGAH DR REK JAMINAN'
        detected, manual, _ = app.process_bank_sources(pd.DataFrame([transaction(narration, 410800)]))
        self.assertEqual(detected.iloc[0]['Kode_8_Digit'], '01010112')
        self.assertEqual(detected.iloc[0]['Bulan'], 'September')
        self.assertEqual(app.determine_payment_method(detected.iloc[0]), 'PEMINDAHBUKUAN')
        self.assertEqual(detected.iloc[0]['_bank_name'], 'Penghuni Contoh')

    def test_multiple_periods_keep_identity_but_never_assign_the_full_amount_to_one_month(self):
        narration = 'PBY SEWA JULI & AGUSTUS 2026 AN PENGHUNI GABUNG HUNIAN B III 03 RSN CIBEUREUM DR REK JAMINAN'
        detected, manual, _ = app.process_bank_sources(pd.DataFrame([transaction(narration, 770000)]))
        eligible, manual = app.filter_incomplete_data(detected, manual, ['2026'])
        self.assertTrue(eligible.empty)
        self.assertEqual(manual.iloc[0]['Kode_8_Digit'], '02020303')
        self.assertEqual(manual.iloc[0]['Credit Transaction'], 770000)
        self.assertEqual(manual.iloc[0]['Bulan'], 'Juli, Agustus')
        self.assertIn('alokasikan manual', manual.iloc[0]['Keterangan'])

    def test_qris_amount_name_or_account_number_cannot_invent_a_housing_identity(self):
        narratives = ['ESB-V001-123456 Payment QRIS MPAN 9360011011000004567 RETRIBUSI',
                      'ESB-V001-123456 Payment QRIS MPAN 1295010101011026 RETRIBUSI',
                      'TRF DARI PENGIRIM CONTOH NOREK 0125422241100',
                      'SETOR TUNAI REFERENSI 990101010188 OKTOBER2026',
                      'MBL PAY VA 1295030299991026']
        for narration in narratives:
            with self.subTest(narration=narration):
                detected, manual, _ = app.process_bank_sources(pd.DataFrame([transaction(narration)]))
                self.assertTrue(detected.empty)
                self.assertEqual(len(manual), 1)
                self.assertFalse(manual.iloc[0]['Kode_8_Digit'])

    def test_debit_or_zero_credit_is_not_promoted_even_with_a_valid_va(self):
        for credit, debit in ((0, 400000), (400000, 400000)):
            detected, manual, _ = app.process_bank_sources(pd.DataFrame([
                transaction('PAY VA 1295030205241026', credit, debit)]))
            self.assertTrue(detected.empty)
            self.assertEqual(len(manual), 1)

    def test_conflicting_coordinates_or_multiple_va_numbers_stay_manual(self):
        for narration in ('PAY VA 1295030205241026 1295020104170726',
                          'SETOR TUNAI 03020524 PBY SEWA OKTOBER 2026 AN A HUNIAN A I 12 RSN CIGUGUR TENGAH'):
            detected, manual, _ = app.process_bank_sources(pd.DataFrame([transaction(narration)]))
            self.assertTrue(detected.empty)
            self.assertIn('Beberapa', manual.iloc[0]['Keterangan'])

    def test_missing_period_is_not_replaced_with_the_bank_posting_month(self):
        detected, manual, _ = app.process_bank_sources(pd.DataFrame([
            transaction('SETOR TUNAI FREDY 03020519 B V 19 LEUWIGAJAH', 392700)]))
        eligible, manual = app.filter_incomplete_data(detected, manual, ['2026'])
        self.assertTrue(eligible.empty)
        self.assertEqual(manual.iloc[0]['Kode_8_Digit'], '03020519')
        self.assertEqual(manual.iloc[0]['Bulan'], '')
        self.assertEqual(manual.iloc[0]['No Hunian'], 19)

    def test_master_name_wins_over_a_bank_preview_name_without_duplicate_columns(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'CIGUGUR'
        for column in ('B', 'C', 'D'):
            sheet[f'{column}3'] = '01'
        sheet['CU3'], sheet['CW3'] = 'Nama Master', 300000
        data = pd.DataFrame([dict(Kode_8_Digit='01010101', Bulan='Oktober', Tahun='2026',
                                  **{'Nama Penghuni': 'Nama Bank', 'Sewa Hunian': 123})])
        with patch('openpyxl.load_workbook', return_value=workbook):
            result = app.extract_from_master_excel(data, {'2026': 'master.xlsx'})
        self.assertTrue(result.columns.is_unique)
        self.assertEqual(result.iloc[0]['Nama Penghuni'], 'Nama Master')
        self.assertEqual(result.iloc[0]['Sewa Hunian'], '300000')

    def test_alternative_narration_header_is_normalized(self):
        data = pd.DataFrame([transaction('SETOR TUNAI 01010101 OKTOBER2026')]).rename(columns={'Narasi': 'NARASI TRANSAKSI'})
        detected, _, _ = app.process_bank_sources(data)
        self.assertEqual(detected.iloc[0]['Kode_8_Digit'], '01010101')

    def test_cash_receipt_matches_only_explicit_identical_unit_period_date_and_amount(self):
        narration = 'PBY SEWA SEPTEMBER 2026 AN PENGHUNI CONTOH HUNIAN A I 12 RSN CIGUGUR TENGAH DR REK JAMINAN'
        receipt = dict(Kode_8_Digit='01010112', Bulan='September', Tahun='2026',
                       Tanggal_Kasda=pd.Timestamp('2026-10-08'), Baris_Kasda=15,
                       Waktu_Bayar_Kasda='12:00:00', Metode_Bayar_Kasda='MANUAL',
                       No_Referensi_Kasda='manual', Nama_Kasda='Penghuni Contoh',
                       Jumlah_Kasda=410800, Denda_Kasda=13300, Keterangan_Pencocokan='')
        for code, matched in (('01010112', True), ('01010113', False)):
            with self.subTest(code=code), patch('app.read_kasda_sewa', return_value=pd.DataFrame([{**receipt, 'Kode_8_Digit': code}])):
                detected, manual, audit = app.process_bank_sources(pd.DataFrame([transaction(narration, 410800)]), 'kasda.xlsx')
            self.assertEqual(detected.iloc[0]['Kode_8_Digit'], '01010112')
            self.assertEqual(audit.iloc[0]['Status_Pencocokan'] == 'Cocok', matched)
            self.assertEqual(len(manual), 0 if matched else 1)

        with patch('app.read_kasda_sewa', return_value=pd.DataFrame([receipt, receipt])):
            _, manual, audit = app.process_bank_sources(pd.DataFrame([transaction(narration, 410800)]), 'kasda.xlsx')
        self.assertEqual(audit['Status_Pencocokan'].eq('Cocok').sum(), 0)
        self.assertEqual(len(manual), 2)

    def test_partial_or_penalty_only_payment_does_not_mark_the_rental_period_paid(self):
        data = pd.DataFrame([dict(ID_Transaksi='BANK-1', Kode_8_Digit='02040218',
            Bulan='Juli', Tahun='2026', **{'Nama Penghuni': 'Penghuni',
            'Tanggal Perjanjian Sewa': '', 'Sewa Hunian': 397500,
            'Sewa Lahan Lantai 1': 60000, 'Credit Transaction': 15810})])
        good, manual = app.filter_unmatched_master(data, data.head(0))
        self.assertTrue(good.empty)
        self.assertEqual(manual.iloc[0]['Credit Transaction'], 15810)
        self.assertIn('tidak boleh menandai lunas', manual.iloc[0]['Keterangan'])
        self.assertEqual(manual.iloc[0]['Rusunawa'], 'Cibeureum')


if __name__ == '__main__':
    unittest.main()
