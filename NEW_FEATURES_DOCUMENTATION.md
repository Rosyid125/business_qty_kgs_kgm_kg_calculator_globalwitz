# Dokumentasi Fitur Baru - Business Quantity Converter

## Overview

Program telah diperbarui untuk menangani case baru pada kolom GSM dan Width yang dapat berisi:

1. **Multiple values separated by comma** (nilai-nilai yang dipisahkan koma)
2. **Operator expressions** pada kolom GSM (>, <, >=, <=)

## Fitur Baru

### 1. Extract First Value (Ambil Nilai Pertama)

Fungsi `extract_first_value()` mengambil nilai pertama dari string yang berisi beberapa nilai terpisah koma.

**Contoh:**

- Input: `"67.80,50,40"` → Output: `"67.80"`
- Input: `"150,200,250"` → Output: `"150"`
- Input: `"100"` → Output: `"100"` (single value tetap)

### 2. Clean GSM Operators (Hapus Operator GSM)

Fungsi `clean_gsm_operators()` menghapus operator matematika dari nilai GSM dan mengambil angka di belakangnya.

**Contoh:**

- Input: `">40"` → Output: `"40"`
- Input: `"<=30"` → Output: `"30"`
- Input: `">=25"` → Output: `"25"`
- Input: `"<50"` → Output: `"50"`
- Input: `"200"` → Output: `"200"` (no operator tetap)

### 3. Combined Processing (Pemrosesan Gabungan)

Untuk kolom GSM, kedua fungsi diterapkan secara berurutan:

**Case 1: Comma-separated dengan operator**

- Input: `"<=30,35,40"`
- Step 1 (extract first): `"<=30"`
- Step 2 (clean operators): `"30"`
- Final numeric: `30.0`

**Case 2: Comma-separated tanpa operator**

- Input: `"67.80,50,40"`
- Step 1 (extract first): `"67.80"`
- Step 2 (clean operators): `"67.80"` (no change)
- Final numeric: `67.80`

## Implementation Details

### Perubahan pada `convert_business_quantity_to_kg()`

**Before (sebelum):**

```python
width = pd.to_numeric(row.get(columns['width'], 0), errors='coerce') or 0
gsm = pd.to_numeric(row.get(columns['gsm'], 0), errors='coerce') or 0
```

**After (sesudah):**

```python
# Process Width: extract first value from comma-separated string
width_raw = row.get(columns['width'], 0) if columns['width'] else 0
width_first = self.extract_first_value(width_raw) if width_raw else None
width = pd.to_numeric(width_first, errors='coerce') or 0 if width_first else 0

# Process GSM: extract first value and remove operators
gsm_raw = row.get(columns['gsm'], 0) if columns['gsm'] else 0
gsm_first = self.extract_first_value(gsm_raw) if gsm_raw else None
gsm_cleaned = self.clean_gsm_operators(gsm_first) if gsm_first else None
gsm = pd.to_numeric(gsm_cleaned, errors='coerce') or 0 if gsm_cleaned else 0
```

## Test Data

### Sample Input Data

| Unit of Weight | Business Quantity | Unit Price (USD) | Width       | GSM          |
| -------------- | ----------------- | ---------------- | ----------- | ------------ |
| KG             | 100               | 15.50            | 150,200,250 | >200,250,300 |
| MTR            | 50                | 25.00            | 120         | <=150,180    |
| MTK            | 25                | 30.00            | 180,220     | >=180        |
| KG             | 200               | 12.75            | 100,150     | <100,120     |
| MTR            | 75                | 40.00            | 160,200,300 | >300,350     |

### Expected Processing Results

**Width Processing:**

- `"150,200,250"` → `150`
- `"120"` → `120`
- `"180,220"` → `180`
- `"100,150"` → `100`
- `"160,200,300"` → `160`

**GSM Processing:**

- `">200,250,300"` → first: `">200"` → cleaned: `"200"` → numeric: `200`
- `"<=150,180"` → first: `"<=150"` → cleaned: `"150"` → numeric: `150`
- `">=180"` → first: `">=180"` → cleaned: `"180"` → numeric: `180`
- `"<100,120"` → first: `"<100"` → cleaned: `"100"` → numeric: `100`
- `">300,350"` → first: `">300"` → cleaned: `"300"` → numeric: `300`

## Error Handling

Fungsi-fungsi baru dilengkapi dengan error handling:

- **Empty/None values**: Mengembalikan `None`
- **Invalid numeric**: `pd.to_numeric(..., errors='coerce')` mengembalikan `NaN` yang dikonversi ke `0`
- **Malformed operators**: Regex pattern menangani spacing dan format yang berbeda

## Backward Compatibility

Fitur baru ini **100% backward compatible**:

- Data dengan single value (tanpa koma) tetap diproses dengan benar
- Data tanpa operator tetap diproses dengan benar
- Konversi existing units tidak terpengaruh

## Testing

File test tersedia:

- `simple_test.py`: Test dasar untuk fungsi baru
- `create_test_data.py`: Membuat sample data Excel untuk testing
- `test_data_new_features.xlsx`: Sample data Excel dengan berbagai case

Jalankan dengan:

```bash
python simple_test.py
python create_test_data.py
```

## Usage Instructions

1. **Prepare Excel file** dengan data yang mengandung comma-separated values atau operators
2. **Place file** di folder `./input/`
3. **Run program** dan pilih file
4. **Map columns** seperti biasa (Width → Width column, GSM → GSM column)
5. **Process data** - program akan otomatis:
   - Mengambil nilai pertama dari Width yang comma-separated
   - Mengambil nilai pertama dari GSM dan menghapus operator
   - Melakukan konversi ke KG seperti biasa

Program akan menghasilkan log yang menunjukkan berapa banyak konversi yang berhasil dan statistik per unit.
