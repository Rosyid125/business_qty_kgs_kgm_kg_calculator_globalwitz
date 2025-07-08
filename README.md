# Business Quantity to KG Converter (Python GUI Version)

Aplikasi GUI untuk mengkonversi Business Quantity ke Kilogram dengan berbagai satuan unit.

## 🚀 Fitur Utama

- **GUI Interface** dengan Tkinter yang user-friendly
- **Multiple Sheet Processing** - Proses beberapa sheet sekaligus
- **Smart Sheet Detection** - Otomatis filter sheet yang valid
- **Real-time Progress** - Monitor proses konversi
- **Error Handling** - Penanganan error yang robust
- **Multiple Unit Support** - Mendukung berbagai satuan berat

## 📋 Satuan yang Didukung

### Direct Conversion (Konversi Langsung - Tanpa parameter tambahan):

#### Kilogram Group:
- **KG, KGS, KGM, K, KILO, KILOS, KILOGRAM, KILOGRAMME** → Basis utama (faktor: 1.0)

#### Gram Group:
- **G, GR, GRM, GRAM, GRAMS, GRAMME, GM, GMS** → ÷ 1000 (gram ke kilogram)

#### Pound Group:
- **LB, LBS, POUND, POUNDS, PND, PNDS, LBM** → × 0.453592 (pounds ke kilogram)

#### Ounce Group:
- **OZ, OUNCE, OUNCES, ONZ** → × 0.0283495 (ounces ke kilogram)

#### Ton Group:
- **TON, TONS, TONNE, TONNES, T** → × 1000 (metric ton ke kilogram)
- **MT, METRICTON, METRICTONS** → × 1000 (metric ton ke kilogram)
- **SHORTTON** → × 907.185 (US short ton ke kilogram)
- **LONGTON** → × 1016.05 (UK long ton ke kilogram)

#### Stone & Imperial Units:
- **ST, STONE, STONES** → × 6.35029 (stone UK ke kilogram)
- **QUINTAL, QUINTALS, Q, QTL** → × 100 (quintal ke kilogram)

#### Precision Units:
- **GRAIN, GRAINS, GRN** → × 0.00006479891 (grain ke kilogram)
- **CARAT, CARATS, CT, CAR** → × 0.0002 (carat ke kilogram)
- **MG, MILLIGRAM, MILLIGRAMS** → × 0.000001 (milligram ke kilogram)
- **UG, MCG, MICROGRAM, MICROGRAMS** → × 0.000000001 (microgram ke kilogram)

#### Additional Imperial Units:
- **DRAM** → × 0.0017718 (dram ke kilogram)
- **SCRUPLE** → × 0.001296 (scruple ke kilogram)
- **PENNYWEIGHT** → × 0.001555 (pennyweight ke kilogram)
- **SLUG** → × 14.5939 (slug ke kilogram)
- **HUNDREDWEIGHT** → × 50.8023 (hundredweight UK ke kilogram)
- **USHUNDREDWEIGHT** → × 45.3592 (hundredweight US ke kilogram)

### Complex Conversion (Butuh Parameter Tambahan):

#### Linear & Area Units:
- **MTR, METER, METRE, M, MTS** → Formula: (Unit Price × 1000) ÷ (Width × GSM)
- **MTK, MTR2, M2, SQM, SQMETER** → Formula: (Unit Price × 1000) ÷ GSM
- **YD, YARD, YARDS, YDS** → Formula: ((Unit Price ÷ 0.9144) × 1000) ÷ (Width × GSM)
- **ROL, ROLL, ROLLS** → Formula: Business Quantity ÷ GSM

### 🎯 Smart Unit Recognition

Aplikasi sekarang mendukung **pengenalan satuan yang sangat fleksibel**:

#### Contoh Variasi yang Didukung:
- **"kg"** = **"KG"** = **"kgs"** = **"K"** = **"kilo"** = **"kilogram"**
- **"lb"** = **"LB"** = **"lbs"** = **"LBS"** = **"pound"** = **"pounds"**
- **"g"** = **"gr"** = **"gram"** = **"grams"** = **"grm"** = **"gms"**
- **"oz"** = **"ounce"** = **"ounces"** = **"onz"**

#### Normalisasi Otomatis:
- ✅ **Case insensitive** - "kg", "KG", "Kg" semua sama
- ✅ **Spasi diabaikan** - "k g", "kg", " kg " semua sama  
- ✅ **Tanda baca diabaikan** - "k.g", "k-g", "kg" semua sama
- ✅ **Variasi spelling** - "kilogram", "kilogramme", "kilo" semua dikenali

#### Statistik Konversi:
Aplikasi akan menampilkan laporan detail tentang:
- Berapa banyak dari setiap unit yang berhasil dikonversi
- Tingkat keberhasilan konversi per unit
- Unit yang tidak dikenali atau gagal dikonversi

## 🛠️ Installation & Setup

### Prerequisites
- Windows 10/11
- Python 3.7+ (Download dari https://python.org)

### Quick Start
1. **Download** semua files ke folder project
2. **Double-click** `start_converter.bat` untuk menjalankan aplikasi
3. **Atau double-click** `run_tests.bat` untuk menjalankan test suite
4. Script akan otomatis install dependencies dan menjalankan aplikasi

### Manual Installation
```bash
# Install dependencies
pip install -r requirements.txt

# Run tests (optional)
python run_tests.bat

# Run application
python business_quantity_converter.py
```

## 📁 Folder Structure
```
project_folder/
├── business_quantity_converter.py  # Main application
├── requirements.txt               # Python dependencies
├── run_converter.bat             # Quick start script
├── input/                        # Put your Excel files here
├── output/                       # Converted files will be saved here
└── README.md                     # This file
```

## 🎯 How to Use

### Step 1: Prepare Excel Files
1. Place your Excel files (.xlsx or .xls) in the `input/` folder
2. Files should have columns for:
   - Unit of Weight (Required)
   - Business Quantity (Required)
   - Unit Price (Optional, needed for MTR/YD/MTK calculations)
   - Width (Optional, needed for MTR/YD calculations)
   - GSM (Optional, needed for complex calculations)

### Step 2: Run Application
1. Double-click `run_converter.bat` or run `python business_quantity_converter.py`
2. The GUI will open

### Step 3: Select Sheet and Map Columns
1. **Select Sheet** from dropdown (this will be the sheet to process)
2. **Map Columns** to the required fields (columns update automatically)

### Step 4: Start Conversion
1. Click **"Start Conversion"** button
2. Monitor progress in the log area
3. Converted file will be saved in `output/` folder with prefix `converted_`

## 🔧 GUI Components

### File Selection
- Lists all Excel files in input folder
- Click "Refresh Files" to update list
- Select file to load its sheets

### Sheet Selection
- **Single Sheet Processing**: Choose one sheet from dropdown to process
- **Dynamic Column Loading**: Columns update automatically when sheet is selected
- **Smart Filtering**: Only shows sheets with valid data
- **Single Workflow**: Select sheet → Map columns → Process

### Column Mapping
- **Auto-Update**: Columns refresh automatically when sheet is changed
- **Unit of Weight**: Required column containing unit codes
- **Business Quantity**: Required column with quantity values
- **Unit Price**: Optional, needed for MTR/YD/MTK calculations
- **Width**: Optional, needed for MTR/YD calculations
- **GSM**: Optional, needed for complex unit calculations

### Process Log
- Real-time feedback during conversion
- Shows progress, errors, and results
- Scrollable text area for long processes

## ⚡ Performance Features

- **Multi-threading** - GUI remains responsive during processing
- **Efficient data handling** - Processes large files quickly
- **Memory optimization** - Handles large datasets efficiently
- **Progress indicators** - Visual feedback for long operations

## 🚨 Error Handling

- **File validation** - Checks file format and accessibility
- **Sheet validation** - Filters out empty or invalid sheets
- **Data validation** - Handles missing or invalid data gracefully
- **User feedback** - Clear error messages and suggestions

## 📊 Output

- **Excel format** - Maintains original formatting
- **New column** - "BUSINESS QUANTITY (KG)" added to each sheet
- **Preserved data** - Original data remains unchanged
- **Multiple sheets** - All selected sheets processed in one file

## 🔍 Troubleshooting

### Common Issues:

1. **"No Excel files found"**
   - Make sure .xlsx or .xls files are in the `input/` folder

2. **"No valid sheets found"**
   - Check if sheets have data
   - Some sheets might be hidden or empty

3. **"Python not found"**
   - Install Python from https://python.org
   - Make sure Python is added to PATH

4. **"Permission denied"**
   - Close Excel if the file is open
   - Check folder permissions

5. **"Module not found"**
   - Run: `pip install -r requirements.txt`

6. **"Unit not recognized"**
   - Check the conversion statistics in the log
   - Unit might have unusual spelling or format
   - Refer to supported units list above

7. **"Low conversion success rate"**
   - Check if required parameters (Unit Price, Width, GSM) are provided for complex units
   - Verify data quality in Business Quantity column
   - Review unit spelling variations

### Unit Recognition Tips:

- **Case doesn't matter**: "kg", "KG", "Kg" all work
- **Spaces are ignored**: "k g", " kg ", "kg" all work  
- **Punctuation is ignored**: "k.g", "k-g", "kg" all work
- **Common abbreviations**: "lb" = "lbs", "g" = "gr" = "gram"
- **Check logs**: The app shows which units were recognized and conversion rates

## 💡 Tips

- **Large files**: Use progress indicator to monitor processing
- **Multiple sheets**: Select all needed sheets at once for batch processing
- **Column mapping**: Required columns must be mapped for conversion to work
- **Backup**: Original files in input folder remain unchanged

## 🧪 Testing & Validation

### Test Files Included:
- **`test_units.py`** - Unit tests for normalization and conversion factors
- **`create_sample_excel.py`** - Generates comprehensive test Excel file

### To Test the Application:
```bash
# Generate sample Excel file with 47 test cases
python create_sample_excel.py

# Run unit tests (100% pass rate expected)
python test_units.py

# Start the main application
python business_quantity_converter.py
```

### Sample Test Data:
The generated test file includes:
- **47 test records** with various unit formats
- **Direct conversions**: KG, G, LBS, OZ, TON, MG, CARAT, STONE, etc.
- **Complex conversions**: MTR, MTK, YD, ROLL with required parameters
- **Spelling variations**: "kg" vs "k g" vs "k.g" vs "kilo"
- **Case variations**: "kg" vs "KG" vs "Kg"
- **4 separate sheets** for organized testing

## 🆕 New Features vs JavaScript Version

✅ **GUI Interface** - No command line needed  
✅ **Visual Progress** - See processing status  
✅ **Multi-threading** - Non-blocking UI  
✅ **Better Error Handling** - User-friendly error messages  
✅ **File Management** - Visual file and sheet selection  
✅ **Real-time Logs** - See what's happening  
✅ **Robust Unit Recognition** - Smart unit detection with 50+ unit variants  
✅ **Comprehensive Weight Units** - Support for all major weight measurement systems  
✅ **Conversion Statistics** - Detailed reporting of conversion success rates  
✅ **Flexible Input** - Case-insensitive, space-tolerant unit recognition  

## 📝 Version History

- **v2.0** - Enhanced unit recognition and comprehensive weight unit support
  - Added 50+ unit variants and aliases
  - Smart normalization system (case-insensitive, space-tolerant)
  - Support for Imperial, Metric, and precision weight units
  - Detailed conversion statistics and reporting
  - Improved error handling and user feedback
- **v1.0** - Initial Python GUI version with all JavaScript features
  - Support for basic unit conversions (K, LBS, etc.)
  - Optimized performance for large files
  - Enhanced user experience with GUI

---

**Made with ❤️ for GlobalWitz by GitHub Copilot**
