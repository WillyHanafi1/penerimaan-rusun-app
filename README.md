# 🏢 Aplikasi Penerimaan Rusun - Streamlit

Aplikasi web untuk memproses data penerimaan rusun dari PDF Bank Statement dan mengintegrasikannya dengan Excel Master.

## 🚀 Fitur Utama

- **Upload PDF Bank Statement**: Ekstraksi otomatis data transaksi dengan algoritma yang sama persis seperti test3.ipynb
- **Pemrosesan SETORTUNAI**: Filter dan ekstraksi data rusun menggunakan regex yang sama persis seperti test3.ipynb
- **Integrasi Master Excel**: Extract data dari file master
- **Kalkulasi Denda**: Otomatis berdasarkan keterlambatan pembayaran
- **Input ke Master**: Update file Excel master dengan backup otomatis
- **Export Multi-Sheet**: File Excel dengan beberapa sheet hasil

## 📋 Alur Kerja

1. **Upload File**: Upload PDF bank statement dan file Excel master
2. **Ekstrak SETORTUNAI**: Pemisahan data rusun dan non-rusun
3. **Filter & Master**: Filter data tidak lengkap dan ekstrak dari master
4. **Input & Export**: Input ke master Excel dan export hasil
5. **Download**: Download file hasil dalam format Excel

## 🛠️ Instalasi

1. **Clone atau download file**:
   ```bash
   # Download file app.py dan requirements.txt
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Jalankan aplikasi**:
   ```bash
   streamlit run app.py
   ```

4. **Akses aplikasi**:
   - Buka browser ke `http://localhost:8501`

## 📁 Struktur File

```
Project/
├── app.py              # Aplikasi Streamlit utama
├── requirements.txt    # Dependencies Python
├── README.md          # Dokumentasi ini
└── temp/              # Folder temporary (auto-created)
    ├── bank_files/    # File PDF bank statement
    ├── Master Data/   # File Excel master
    └── output/        # File hasil export
```

## 📊 Format File Input

### PDF Bank Statement
- Format: PDF dengan tabel transaksi
- Kolom yang diperlukan: Posting Date, Effective Date, Narasi, Credit Transaction, dll.
- Transaksi harus mengandung kata "SETORTUNAI" untuk data rusun
- **Catatan**: Semua spasi di kolom Narasi akan dihilangkan otomatis saat ekstraksi

### Excel Master
- Format: .xlsx dengan sheet CIGUGUR, MELONG, LG
- Kolom H: Kode 8 digit rusun
- Kolom per bulan untuk data nama penghuni, tanggal perjanjian, sewa

## 📤 Format File Output

### Sheet 'Status_Input'
- Status proses input (Berhasil/Skip/Gagal)
- Keterangan detail
- Nilai denda yang diinput

### Sheet 'Cek Manual'
- Data NON-RUSUN
- Data SETORTUNAI tidak lengkap

### Sheet 'Kasda'
- Data semua transaksi untuk kasda
- Format tanggal: dd-mmm-yy
- Format angka dengan titik pemisah ribuan

### Sheet 'Denda'
- Data transaksi dengan denda > 0
- Format sama dengan sheet Kasda

## ⚙️ Konfigurasi

### Mapping Rusunawa
- 01: Cigugur Tengah
- 02: Cibeureum  
- 03: Leuwigajah

### Mapping Gedung
- 01: A, 02: B, 03: C, 04: D

### Mapping Lantai
- 01: I, 02: II, 03: III, 04: IV

### Kalkulasi Denda
- Rate: 2% per bulan keterlambatan
- Basis: Sewa hunian
- Mulai dihitung dari bulan ke-2 keterlambatan

## 🔧 Troubleshooting

### Error "Import streamlit could not be resolved"
```bash
pip install streamlit
```

### Error "Columns must be same length as key"
- Error ini terjadi saat ekstraksi SETORTUNAI
- Biasanya karena data tahun tidak konsisten (string vs integer)
- Sudah diperbaiki dengan safe type conversion

### Error PDF tidak bisa dibaca
- Pastikan PDF berisi tabel yang dapat diekstrak
- Coba PDF yang tidak di-scan (native PDF)

### Error Excel tidak bisa dibuka
- Pastikan file Excel tidak sedang dibuka di aplikasi lain
- Cek format file (.xlsx)

### Memory error saat memproses file besar
- Tutup aplikasi lain
- Gunakan file yang lebih kecil
- Restart aplikasi Streamlit

### Error ekstraksi data tidak lengkap
- Cek format narasi di PDF bank statement
- Pastikan mengandung kata "SETORTUNAI"
- Periksa pola kode 8 digit dan tahun dalam narasi

## 📞 Support

Jika mengalami masalah:
1. Cek console/terminal untuk error message
2. Pastikan semua dependencies terinstall
3. Cek format file input sesuai spesifikasi

## 🆕 Update Log

- **v1.0**: Implementasi awal dari notebook test3.ipynb
- Fitur lengkap: PDF processing, Excel integration, multi-sheet export
- UI Streamlit dengan 5 tahapan proses
- Backup otomatis file master
- Format numeric dengan titik sebagai pemisah ribuan

## 📝 Catatan

- File temporary akan dibuat di folder sistem temp
- Backup file Excel dibuat otomatis sebelum modifikasi
- Session state menyimpan data antar tahapan proses
- Aplikasi mendukung file master untuk tahun 2024 dan 2025
