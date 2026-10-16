from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import openpyxl
import pandas as pd

import app
from master_formula_cache import NS, save_master


class MasterOutputTests(unittest.TestCase):
    def test_download_omits_technical_columns_after_matching_status_and_keeps_amounts(self):
        hidden = {
            'Urutan_Rekening_Koran', 'ID_Transaksi', 'Narasi_Clean', 'Has_SETORTUNAI',
            'Waktu_Bayar_Kasda', 'No_Referensi_Kasda', 'Nama_Kasda', 'Sheet_Kasda',
            'Tanggal_Kasda', 'Tanggal_Bayar_Kasda', 'Biaya_Lain_Kasda',
            'Nomor_Dokumen_Kasda', 'Metode_Pencocokan',
        }
        rows = []
        for index in (1, 2):
            rows.append({**dict.fromkeys(hidden, 'data teknis'),
                'ID_Transaksi': f'BANK-{index}', 'Kode_8_Digit': '01010101',
                'Bulan': 'Oktober', 'Tahun': '2026', 'Posting Date': '14 Oct 2026',
                'Narasi': 'Baris pertama\nBaris kedua',
                'Nama Penghuni': 'Penghuni dengan nama panjang untuk pemeriksaan lebar kolom',
                'Sewa Hunian': 100000, 'Sewa Lahan Lantai 1': 0,
                'Tanggal Perjanjian Sewa': '01/01/2026', 'Denda': 5000,
                'Credit Transaction': 105000, 'Balance': 210000})
        data = pd.DataFrame(rows)
        original = data.copy(deep=True)
        results = {
            'success_details': [dict(ID_Transaksi='BANK-1', Kode_8_Digit='01010101',
                Bulan='Oktober', Sheet='CIGUGUR', Row=3, Denda=5000)],
            'skipped_details': [dict(ID_Transaksi='BANK-2', Kode_8_Digit='01010101',
                Bulan='Oktober', Reason='Sudah terisi')],
        }
        for audit in (None, data):
            with self.subTest(audit=audit is not None), tempfile.TemporaryDirectory() as folder:
                report = app.create_export_excel(results, data, data, data.head(1), audit, folder)
                workbook = openpyxl.load_workbook(report)
                self.assertEqual(workbook.sheetnames, ['Status_Input', 'Cek Manual', 'Kasda', 'Denda']
                                 + (['Pencocokan Kasda'] if audit is not None else []))
                for sheet in workbook:
                    headers = [cell.value for cell in sheet[1]]
                    self.assertFalse(hidden.intersection(headers))
                    self.assertEqual(sheet.freeze_panes, 'A2')
                    self.assertEqual(sheet.auto_filter.ref, sheet.dimensions)
                    self.assertTrue(all(dimension.height == 18 for dimension in sheet.row_dimensions.values()))
                    self.assertTrue(all(not cell.alignment.wrap_text for row in sheet for cell in row))
                status = workbook['Status_Input']
                headers = {cell.value: cell.column for cell in status[1]}
                self.assertEqual([status.cell(row, headers['Status_Input']).value for row in (2, 3)],
                                 ['Berhasil Input', 'Skip - Sudah Terisi'])
                for row in (2, 3):
                    self.assertEqual(status.cell(row, headers['Credit Transaction']).value, 105000)
                    self.assertEqual(status.cell(row, headers['Credit Transaction']).number_format, '#,##0')
                    self.assertEqual(status.cell(row, headers['Kode_8_Digit']).value, '01010101')
                    self.assertEqual(status.cell(row, headers['Kode_8_Digit']).number_format, '@')
                    self.assertEqual(status.cell(row, headers['Narasi']).value, 'Baris pertama Baris kedua')
                name_column = status.cell(1, headers['Nama Penghuni']).column_letter
                self.assertGreater(status.column_dimensions[name_column].width, 18)
        pd.testing.assert_frame_equal(data, original)

    def test_blank_or_zero_tariffs_do_not_turn_the_whole_payment_into_a_penalty(self):
        data = pd.DataFrame([
            {'Kode_8_Digit': f'0101010{i}', 'Bulan': 'Oktober', 'Tahun': '2026',
             'Nama Penghuni': 'Contoh', 'Tanggal Perjanjian Sewa': '',
             'Sewa Hunian': rent, 'Sewa Lahan Lantai 1': '', 'Credit Transaction': 310000}
            for i, rent in enumerate(('', 0, 300000), 1)
        ])
        eligible, manual = app.filter_unmatched_master(data, data.head(0))
        self.assertEqual(eligible['Kode_8_Digit'].tolist(), ['01010103'])
        self.assertEqual(len(manual), 2)
        self.assertTrue(manual['Keterangan'].str.contains('Tarif sewa master kosong/nol').all())
        self.assertEqual(app.calculate_denda(eligible).iloc[0]['Denda'], 10000)

    def test_fixed_tariff_formula_cache_survives_but_changed_total_recalculates(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'master.xlsx'
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet['K3'] = '=280000+90000'
            sheet['K4'] = '=280000+90000'
            sheet['M3'] = 0
            sheet['N3'] = '=K3+M3'
            workbook.save(path)
            # Simulate the cached results written by Excel in a real input master.
            buffer = BytesIO()
            with ZipFile(path) as source, ZipFile(buffer, 'w') as out:
                for entry in source.infolist():
                    content = source.read(entry.filename)
                    if entry.filename == 'xl/worksheets/sheet1.xml':
                        root = ET.fromstring(content)
                        for cell in root.iter(NS + 'c'):
                            if cell.get('r') in ('K3', 'K4', 'N3'):
                                cell.find(NS + 'v').text = '370000'
                            if cell.get('r') in ('K3', 'K4'):
                                formula = cell.find(NS + 'f')
                                formula.set('t', 'shared')
                                formula.set('si', '0')
                                if cell.get('r')=='K3':
                                    formula.set('ref', 'K3:K4')
                                else:
                                    formula.text = None
                        content = ET.tostring(root)
                    out.writestr(entry, content)
            path.write_bytes(buffer.getvalue())
            workbook = openpyxl.load_workbook(path)
            workbook.active['M3'] = 10000
            save_master(workbook, path)
            values = openpyxl.load_workbook(path, data_only=True)
            self.assertEqual(values.active['K3'].value, 370000)
            self.assertEqual(values.active['K4'].value, 370000)
            self.assertEqual(values.active['M3'].value, 10000)
            self.assertIsNone(values.active['N3'].value)  # no stale cached 370000
            formulas = openpyxl.load_workbook(path)
            self.assertEqual(formulas.active['K3'].value, '=280000+90000')
            self.assertEqual(formulas.active['N3'].value, '=K3+M3')

    def test_all_uploaded_years_are_returned_and_repeated_input_is_skipped(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            masters = {}
            for year in ('2025', '2026'):
                path = root / f'Master_{year}.xlsx'
                workbook = openpyxl.Workbook()
                workbook.active.title = 'CIGUGUR'
                for column in ('B', 'C', 'D'):
                    workbook.active[f'{column}3'] = '01'
                workbook.save(path)
                masters[year] = str(path)
            source_contents = {y: Path(p).read_bytes() for y, p in masters.items()}
            data = pd.DataFrame([dict(ID_Transaksi='BANK-1', Kode_8_Digit='01010101',
                Bulan='Oktober', Tahun='2026', **{'Posting Date': '14 Oct 2026', 'Denda': 0})])
            _, results = app.input_to_excel_master(data.copy(), masters, str(root / 'first'))
            self.assertEqual(results['success'], 1)
            self.assertEqual(set(results['backup_files']), {'2025', '2026'})
            self.assertEqual(Path(results['backup_files']['2025']).read_bytes(), source_contents['2025'])
            for year, path in masters.items():
                self.assertEqual(Path(path).read_bytes(), source_contents[year])
            _, repeated = app.input_to_excel_master(data.copy(), results['backup_files'], str(root / 'second'))
            self.assertEqual((repeated['success'], repeated['skipped'], repeated['failed']), (0, 1, 0))
            self.assertEqual(Path(repeated['backup_files']['2026']).read_bytes(),
                             Path(results['backup_files']['2026']).read_bytes())


if __name__ == '__main__':
    unittest.main()
