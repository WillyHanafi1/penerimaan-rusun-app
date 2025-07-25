# Optimasi Performa Excel Processing

## Masalah Sebelumnya

Pada fungsi `extract_from_master_excel` dan `input_to_excel_master`, operasi baca/tulis file Excel (openpyxl.load_workbook dan workbook.save) dilakukan di dalam perulangan (loop). Ini berarti jika ada 50 transaksi yang harus diproses, file Excel akan dibuka dan ditutup (atau disimpan) sebanyak 50 kali.

### Bottleneck Performa:
- **extract_from_master_excel**: `openpyxl.load_workbook()` dipanggil untuk setiap baris data
- **input_to_excel_master**: `workbook.save()` dipanggil untuk setiap transaksi
- Operasi I/O (membaca/menulis ke disk) sangat lambat

## Solusi Optimasi

### 1. extract_from_master_excel (Optimasi Baca)
**Sebelum:**
```python
def extract_data_excel(row):
    # ... validasi ...
    workbook = openpyxl.load_workbook(excel_file, data_only=True)  # ← DIBUKA SETIAP KALI
    worksheet = workbook[sheet_name]
    # ... ekstrak data ...
```

**Sesudah:**
```python
# OPTIMIZATION: Load workbooks once outside the loop
workbooks = {}
worksheets_cache = {}

# Load all workbooks at the beginning
for year, excel_file in master_files.items():
    workbooks[year] = openpyxl.load_workbook(excel_file, data_only=True)  # ← DIBUKA SEKALI SAJA
    worksheets_cache[year] = {}
    for sheet_name in ['CIGUGUR', 'MELONG', 'LG ']:
        if sheet_name in workbooks[year].sheetnames:
            worksheets_cache[year][sheet_name] = workbooks[year][sheet_name]

def extract_data_excel_optimized(row):
    # ... validasi ...
    worksheet = worksheets_cache[tahun][sheet_name]  # ← GUNAKAN CACHE
    # ... ekstrak data ...
```

### 2. input_to_excel_master (Optimasi Tulis)
**Sebelum:**
```python
for _, row in group_data.iterrows():
    # ... proses data ...
    worksheet[c1_addr] = p_date
    worksheet[c2_addr] = p_date  
    worksheet[c3_addr] = denda
    # ... apply formatting ...
    workbook.save(backup_path)  # ← SIMPAN SETIAP KALI
```

**Sesudah:**
```python
# OPTIMIZATION: Create backup and load workbook ONCE per year
backup_path = create_backup_file(excel_file)
workbook = openpyxl.load_workbook(backup_path)  # ← BUKA SEKALI

# Cache all worksheets to avoid repeated sheet access
worksheets_cache = {}
for sheet_name in ['CIGUGUR', 'MELONG', 'LG ']:
    if sheet_name in workbook.sheetnames:
        worksheets_cache[sheet_name] = workbook[sheet_name]

for _, row in group_data.iterrows():
    # ... proses data ...
    worksheet = worksheets_cache[sheet_name]  # ← GUNAKAN CACHE
    worksheet[c1_addr] = p_date
    worksheet[c2_addr] = p_date
    worksheet[c3_addr] = denda
    # ... apply formatting in memory ...

# OPTIMIZATION: Save workbook ONCE after all changes are done
workbook.save(backup_path)  # ← SIMPAN SEKALI SAJA
```

## Estimasi Peningkatan Performa

### Skenario: 50 Transaksi
**Sebelum Optimasi:**
- extract_from_master_excel: 50x load workbook = ~50-100 detik
- input_to_excel_master: 50x save workbook = ~50-100 detik
- **Total: ~100-200 detik**

**Setelah Optimasi:**
- extract_from_master_excel: 2x load workbook (2024+2025) = ~2-4 detik
- input_to_excel_master: 2x save workbook (2024+2025) = ~2-4 detik  
- **Total: ~4-8 detik**

**Peningkatan: 25-50x lebih cepat!**

## Manfaat Tambahan

1. **Worksheet Caching**: Akses worksheet tidak perlu mencari by name setiap kali
2. **Memory Efficiency**: Workbook tetap di memory selama proses berlangsung
3. **Error Reduction**: Mengurangi kemungkinan error karena file I/O berulang
4. **User Experience**: Aplikasi terasa jauh lebih responsif

## Implementation Notes

- Optimasi dilakukan dengan tetap mempertahankan logic bisnis yang sama
- Backward compatibility terjaga
- Error handling tetap robust
- Cache management otomatis per year/sheet
