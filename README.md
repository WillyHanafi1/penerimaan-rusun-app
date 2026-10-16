# 🏢 Aplikasi Penerimaan Rusun - Streamlit

Aplikasi web untuk memproses data penerimaan rusun dari PDF Bank Statement dan mengintegrasikannya dengan Excel Master.

## Upload Kasda Sewa (opsional)

Di bawah upload PDF terdapat upload **Kasda Sewa**, misalnya file `rekap-pendapatan-2026-09-04 (2).xlsx`.
Pilih tahun master yang akan diproses, unggah master pada slot tahunnya, lalu jalankan proses otomatis.

- Tanpa Kasda, nomor VA rusun dan narasi yang menyebut hunian/periode secara jelas tetap dapat dideteksi. Kode hunian tidak diambil dari potongan nomor rekening panjang, nama pengirim, atau nominal yang kebetulan sama. QRIS tanpa identitas hunian tetap memerlukan Kasda.
- Narasi pemindahbukuan jaminan seperti `PBY SEWA ... HUNIAN ... RSN ...` dibaca berdasarkan lokasi dan periode tertulis. Pembayaran beberapa bulan menampilkan identitasnya di `Cek Manual`; nominalnya tidak dibagi rata atau dimasukkan seluruhnya ke satu bulan.
- Pembayaran tunai/pemindahbukuan hanya ditautkan ke Kasda manual jika hunian, periode, tanggal, dan nominal sama persis serta pasangannya tunggal. Kasda manual tidak memakai pencocokan urutan waktu.
- Lokasi hunian tetap ditampilkan untuk pembayaran yang belum dapat diinput. Nama dan tarif untuk input mengikuti master yang diunggah; nama pada preview bank tidak mengganti nama master.
- Nama penghuni yang kosong tidak menghalangi pencatatan pembayaran. Nama tetap kosong sampai petugas mengisinya berdasarkan kontrak; aplikasi tidak mengisi nama master dari narasi bank atau Kasda.
- Jika tarif master juga kosong, pembayaran tetap dicatat memakai rincian sewa, parkir, dan denda Kasda yang cocok persis lewat referensi, atau hunian/periode untuk pembayaran manual. Jumlah seluruh komponen, termasuk biaya lain, harus sama dengan kredit bank. Nilai sewa Kasda hanya tampil pada laporan; nama, tanggal kontrak, dan tarif di master tetap kosong untuk dilengkapi petugas. Pasangan berdasarkan urutan waktu saja atau rincian yang belum lengkap tetap masuk `Cek Manual`.
- Pembayaran di bawah jumlah tarif master, termasuk pembayaran sebagian atau denda/parkir saja, masuk `Cek Manual` agar tidak menandai sewa bulanan sebagai lunas. Nama/tarif bulan yang kosong pada master tidak diisi dari bulan lain secara otomatis.

- Ekspor harian Inawa terbaru dari `/admin/reports` berisi **satu sheet untuk tanggal yang dipilih**, misalnya `kasda-sewa-2026-10-14.xlsx` hanya memiliki sheet `14-10-2026`. Unggah file ini bersama rekening koran tanggal yang sama dan master sesuai tahun sewa; sheet tanggal sebelumnya tidak diperlukan. Workbook lama dengan beberapa tanggal tetap didukung: aplikasi memilih **sheet sesuai tanggal rekening koran**, bukan sheet pertama/terakhir. Tanggal lain tidak ikut diproses. Sheet yang tidak tersedia atau header rusak menghasilkan pesan jelas, sedangkan sheet kosong yang valid tetap dapat diproses.
- Format lama dengan sheet `Sewa` tetap didukung. Judul kelompok rusun menentukan lokasi. Baris subtotal dan sheet `Denda` tidak dihitung lagi sebagai transaksi.
- Pada template harian Inawa, header gabungan `NOMOR HUNIAN` memetakan gedung/lantai/unit dari tiga kolom detail. `KANAL BAYAR` dibaca sebagai metode bayar dan `WAKTU BAYAR (WIB)` sebagai waktu pembayaran. Referensi lengkap, nomor SKRD/STRD, biaya lain, sheet sumber, dan baris sumber dibawa ke audit pencocokan.
- Pencocokan rekening koran memakai `Tanggal Setor` (tanggal pembayaran). `Tanggal Kasda` disimpan terpisah karena bisa berbeda hari. Untuk format lama tanpa `Tanggal Setor`, tanggal pencocokan mengikuti `Tanggal Kasda`.
- Nomor referensi QRIS (enam karakter terakhir) atau VA dicocokkan dengan rekening koran, beserta tanggal dan nominal pembayaran.
- Jika referensi belum cocok, `WAKTU BAYAR` diurutkan dari paling awal dan dicocokkan dengan urutan pembayaran di PDF. Baris dengan nominal berbeda dilewati, kemudian pencocokan dilanjutkan ke baris berikutnya. Jumlah baris kedua sumber boleh berbeda. Tanggal/metode harus sesuai, pasangan referensi tetap menjadi batas urutan, dan pasangan ambigu tidak dipaksakan.
- Narasi `Payment QRIS ... MPAN ... RETRIBUSI` dibedakan dari pelimpahan yang hanya menyebut QRIS. VA mendukung kanal TLR maupun MBL. Transaksi debit dan transfer lain tidak memakai pasangan urutan Kasda.
- Transaksi bank yang tidak cocok dan penerimaan Kasda tanpa pasangan masuk `Cek Manual`, disertai alasan serta jenis sumber. Nominal sumber Kasda disimpan di `Jumlah_Kasda`, tanpa menciptakan kredit bank tambahan.
- Baris yang gagal menemukan kode/periode pada master dipindahkan ke `Cek Manual` sebelum perhitungan denda. Semua baris yang belum terdeteksi tersebut tidak masuk `Status_Input`, Kasda/Denda otomatis, atau input master.
- Tarif master yang kosong/nol juga masuk `Cek Manual`; seluruh pembayaran tidak boleh otomatis dianggap denda. Lengkapi master lalu proses ulang transaksi tersebut.
- Kode hunian dibentuk dari kelompok rusun, GD, LT, dan UNIT. Bulan/tahun sewa diambil dari kolom `BULAN`, bukan tahun anggaran atau tanggal pembayaran.
- Transaksi yang cocok masuk `Status_Input` dan mengikuti aturan input master yang sama. Sel tujuan yang sudah terisi tetap berstatus `Skip - Sudah Terisi`.
- Pada baris baru yang diinput ke Master 2025/2026, kolom `KETERANGAN` otomatis diisi `QRIS` atau `VA`. Pewarnaan pelunasan disesuaikan: `#C5D9F1` (biru muda) untuk QRIS dan `#D9B5CA` (ungu muda) untuk VA.
- Sheet tambahan `Pencocokan Kasda` mencatat semua penerimaan Kasda, nominal, status pencocokan, serta alasan bila belum cocok. Kolom teknis pencocokan dipakai di dalam proses dan tidak ditampilkan pada unduhan laporan.
- Nama penghuni dan tarif mengikuti master sesuai periode. Denda dihitung dari selisih pembayaran dan tarif master; khusus tarif kosong dengan rincian Kasda terverifikasi, laporan memakai sewa/parkir dan denda Kasda tanpa mengubah tarif master. `Denda_Kasda` disertakan sebagai pembanding.
- Kedua master yang diunggah tetap tersedia untuk diunduh, termasuk tahun tanpa penerimaan baru. Master tanpa perubahan disalin identik. Hasil rumus tarif tetap seperti `=280000+90000` dipertahankan ketika master disimpan agar masih terbaca saat digunakan lagi untuk proses harian; rumus total tetap dihitung ulang oleh Excel. Proses ulang pada master hasil melewati sel yang sudah terisi.

Untuk memproses file lokal tanpa browser, gunakan `python scripts/process_files.py --help`.
Hasil ditulis ke folder baru; checksum file sumber diperiksa agar master dan rekening koran asli tetap utuh.

Format unduhan laporan menghapus kolom teknis `Urutan_Rekening_Koran`, `ID_Transaksi`,
`Narasi_Clean`, `Has_SETORTUNAI`, `Waktu_Bayar_Kasda`, `No_Referensi_Kasda`,
`Nama_Kasda`, `Sheet_Kasda`, `Tanggal_Kasda`, `Tanggal_Bayar_Kasda`,
`Biaya_Lain_Kasda`, `Nomor_Dokumen_Kasda`, `Metode_Pencocokan`, `Jumlah_Kasda`,
`Sumber_Deteksi`, dan `Baris_Kasda` dari seluruh sheet.
Data tersebut tetap dipakai di dalam proses pencocokan. Semua baris laporan memiliki
tinggi 18, tanpa wrap text, dengan lebar kolom menyesuaikan isi.
Pada sheet `Status_Input`, sukses biasa dengan keterangan
`Data berhasil diinput ke <sheet> row <nomor>` tetap memakai warna normal.
Seluruh baris dengan catatan tambahan (termasuk rincian memakai Kasda), status
dilewati, gagal, atau belum diproses diberi warna `#F2DCDB` agar mudah diperiksa.
Pemisah baris di narasi unduhan diganti spasi agar seluruh teks tampil dalam satu baris.

Sheet `Status_Input` berisi tepat 22 kolom dengan urutan: `No`, `Posting Date`,
`Effective Date`, `Narasi`, `Debit Transaction`, `Credit Transaction`, `Balance`,
`Kode_8_Digit`, `Rusunawa`, `Gedung`, `Lantai`, `No Hunian`, `Nama Penghuni`,
`Tanggal Perjanjian Sewa`, `Sewa Hunian`, `Sewa Lahan Lantai 1`, `Denda`,
`Bulan`, `Tahun`, `Kanal Bayar`, `Status_Input`, dan `Keterangan_Input`.
Kanal bayar memakai penentuan metode yang sama dengan input master, misalnya `QRIS` atau `VA`.

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
