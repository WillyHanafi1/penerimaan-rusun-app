"""Run the same import stages as the UI on local files, writing to a new folder."""

import argparse
from datetime import datetime
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app


def process_files(bank_path, master_files, kasda_path, years, output_dir):
    sources = [Path(bank_path), *(Path(p) for p in master_files.values())]
    if kasda_path:
        sources.append(Path(kasda_path))
    hashes = {str(p.resolve()): hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}
    output_dir = Path(output_dir).resolve()
    # A fresh destination prevents accidental replacement of earlier results.
    output_dir.mkdir(parents=True, exist_ok=False)
    bank = app.process_pdf(str(bank_path))
    if bank.empty:
        raise ValueError('Tidak ada transaksi yang dapat dibaca dari PDF.')
    detected, manual, audit = app.process_bank_sources(bank, str(kasda_path) if kasda_path else None)
    eligible, manual = app.filter_incomplete_data(detected, manual, years)
    manual_before_master = len(manual)
    extracted = app.extract_from_master_excel(eligible, master_files)
    matched_master, manual = app.filter_unmatched_master(extracted, manual)
    final = app.calculate_denda(matched_master)
    valid, results = app.input_to_excel_master(final, master_files, str(output_dir))
    report = app.create_export_excel(results, valid, final, manual, audit, str(output_dir))
    unchanged = all(hashlib.sha256(Path(p).read_bytes()).hexdigest() == digest for p, digest in hashes.items())
    summary = {
        'bank_rows': len(bank), 'detected': len(detected), 'eligible': len(eligible),
        'manual_before_master': manual_before_master, 'manual_rows': len(manual),
        'years': eligible['Tahun'].value_counts().to_dict(),
        'manual_bank_rows': int(manual['ID_Transaksi'].fillna('').ne('').sum()),
        'manual_kasda_rows': int(manual['ID_Transaksi'].fillna('').eq('').sum()),
        'unmatched_master': len(extracted) - len(matched_master), 'status_rows': len(valid),
        'input': results, 'kasda_matched': int((audit['Status_Pencocokan'] == 'Cocok').sum()) if audit is not None else 0,
        'kasda_sheets': audit['Sheet_Kasda'].dropna().unique().tolist() if audit is not None and 'Sheet_Kasda' in audit else [],
        'matching_methods': audit['Metode_Pencocokan'].value_counts().to_dict() if audit is not None else {},
        'bank_credit_total': float(bank['Credit Transaction'].apply(app.to_numeric_safe).sum()),
        'bank_debit_total': float(bank['Debit Transaction'].apply(app.to_numeric_safe).sum()),
        'manual_bank_credit_total': float(manual.loc[manual['ID_Transaksi'].fillna('').ne(''), 'Credit Transaction'].apply(app.to_numeric_safe).sum()),
        'payment_total': float(final['Credit Transaction'].apply(app.to_numeric_safe).sum()),
        'denda_total': float(final['Denda'].sum()), 'report': report,
        'source_sha256': hashes, 'sources_unchanged': unchanged,
    }
    (output_dir / 'ringkasan_proses.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    if not unchanged:
        raise RuntimeError('Pemeriksaan integritas file sumber gagal.')
    return summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--bank', required=True, type=Path)
    parser.add_argument('--kasda', type=Path)
    parser.add_argument('--master', action='append', required=True, metavar='YEAR=PATH')
    parser.add_argument('--years', nargs='+')
    parser.add_argument('--output', type=Path, default=Path('outputs') / datetime.now().strftime('run_%Y%m%d_%H%M%S'))
    args = parser.parse_args()
    masters = {}
    for item in args.master:
        year, separator, path = item.partition('=')
        if not separator or not year.isdigit():
            parser.error('--master harus berformat YEAR=PATH')
        masters[year] = str(Path(path).resolve())
    summary = process_files(args.bank, masters, args.kasda, args.years or list(masters), args.output)
    print(json.dumps({k: v for k, v in summary.items() if k not in ('source_sha256', 'input')}, ensure_ascii=False, indent=2))
    print(json.dumps({k: summary['input'][k] for k in ('success', 'skipped', 'failed', 'errors')}, ensure_ascii=False))


if __name__ == '__main__':
    main()
