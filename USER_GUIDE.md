# 🏢 Aplikasi Penerimaan Rusun - One-Click Automation

## 🚀 Cara Penggunaan (Super Mudah!)

### 1. **Upload File** 📤
- **Wajib**: Upload file PDF Bank Statement  
- **Opsional**: Upload Master Excel 2024 dan/atau 2025

### 2. **Proses Otomatis** ⚡
- Klik tombol **"PROSES OTOMATIS - SEMUA LANGKAH"**
- Tunggu 30-60 detik (ada progress bar)

### 3. **Download Hasil** 📥
- Download **Laporan Status** (Excel dengan 4 sheet)
- Download **Master Excel Terupdate** (jika Master diupload)

**Selesai!** ✅

---

## 📋 File yang Diperlukan

### 📄 **PDF Bank Statement** (Wajib)
- File PDF hasil download dari internet banking
- Berisi data transaksi yang akan diproses
- Format: `.pdf`

### 📊 **Master Excel** (Opsional)
- **Master Excel 2024**: Untuk data tahun 2024
- **Master Excel 2025**: Untuk data tahun 2025  
- Format: `.xlsx`
- **Catatan**: Jika tidak upload Master Excel, sistem tetap bisa memproses data PDF

---

## 📊 Hasil yang Didapat

### 1. **File Laporan Status** 
Excel dengan 4 sheet:
- **Status_Input**: Status proses input (Berhasil/Skip/Gagal)
- **Cek Manual**: Data NON-RUSUN yang perlu dicek manual
- **Kasda**: Data semua transaksi untuk kasda
- **Denda**: Data khusus dengan denda > 0

### 2. **Master Excel Terupdate** (jika Master diupload)
- File Master asli + data baru yang sudah diinput
- Highlight hijau pada data yang baru diinput
- Backup timestamp untuk tracking

---

## ⚡ Proses Otomatis (Background)

Sistem akan otomatis menjalankan 6 langkah:

1. **📄 Ekstraksi PDF** - Ambil data transaksi dari PDF
2. **🔍 Ekstraksi SETORTUNAI** - Pisahkan data rusun vs non-rusun  
3. **🛠️ Filter Data** - Filter data lengkap dan tahun didukung
4. **📊 Ekstrak Master** - Ambil data dari Master Excel
5. **💰 Kalkulasi Denda** - Hitung denda otomatis
6. **📝 Input & Export** - Input ke Master + buat laporan

**Progress ditampilkan real-time dengan progress bar!**

---

## 🎯 Keunggulan

- ✅ **One-Click**: Hanya 1 klik untuk semua proses
- ✅ **Super Cepat**: 25-50x lebih cepat berkat optimasi
- ✅ **User-Friendly**: Interface yang sangat mudah
- ✅ **Mobile Ready**: Bisa digunakan di HP/tablet
- ✅ **Error Safe**: Handle error dengan graceful
- ✅ **Production Ready**: Siap untuk penggunaan sehari-hari

---

## 🔧 Persyaratan Sistem

- **Browser**: Chrome, Firefox, Safari, Edge (versi terbaru)
- **Internet**: Untuk mengakses aplikasi Streamlit
- **File Size**: PDF maksimal 50MB, Excel maksimal 100MB
- **Format**: PDF untuk bank statement, XLSX untuk Master Excel

---

## 📱 Tips Penggunaan

1. **Siapkan File**: Pastikan PDF dan Master Excel siap sebelum upload
2. **Koneksi Stabil**: Pastikan internet stabil selama proses
3. **Tunggu Selesai**: Jangan tutup browser saat proses berjalan
4. **Download Langsung**: Download hasil segera setelah selesai
5. **Backup**: Simpan hasil download untuk arsip

---

## 🆘 Troubleshooting

### ❌ "Tidak ada data yang bisa diekstrak dari PDF"
- **Solusi**: Pastikan PDF berisi tabel data transaksi yang valid
- Coba dengan PDF bank statement yang berbeda

### ❌ "Error dalam proses otomatis"  
- **Solusi**: Refresh halaman dan coba lagi
- Pastikan file tidak corrupt
- Cek ukuran file tidak terlalu besar

### ⚠️ "Tidak ada file Master yang diupdate"
- **Informasi**: Normal jika tidak upload Master Excel
- File laporan status tetap tersedia untuk download

### 📊 Data tidak sesuai ekspektasi
- Cek sheet "Cek Manual" untuk data yang perlu validasi
- Review sheet "Status_Input" untuk detail proses

---

## 🎉 Selamat Menggunakan!

Aplikasi ini dirancang untuk memudahkan pekerjaan Anda. Jika ada pertanyaan atau saran, jangan ragu untuk memberikan feedback!

**Happy Processing!** 🚀
