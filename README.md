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

### Direct Conversion (Tanpa parameter tambahan):
- **KG, KGM, KGS, K** → Langsung copy nilai
- **GRM, GR** → Kalikan 1000 (gram ke kilogram)
- **LBS** → Kalikan 0.453592 (pounds ke kilogram)

### Complex Conversion (Butuh Unit Price, Width, GSM):
- **MTR** → Formula: (Unit Price × 1000) ÷ (Width × GSM)
- **MTK, MTR2** → Formula: (Unit Price × 1000) ÷ GSM
- **YD** → Formula: ((Unit Price ÷ 0.9144) × 1000) ÷ (Width × GSM)
- **ROL, ROLL** → Formula: Business Quantity ÷ GSM

## 🛠️ Installation & Setup

### Prerequisites
- Windows 10/11
- Python 3.7+ (Download dari https://python.org)

### Quick Start
1. **Download** semua files ke folder project
2. **Double-click** `run_converter.bat`
3. Script akan otomatis install dependencies dan menjalankan aplikasi

### Manual Installation
```bash
# Install dependencies
pip install -r requirements.txt

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

## 💡 Tips

- **Large files**: Use progress indicator to monitor processing
- **Multiple sheets**: Select all needed sheets at once for batch processing
- **Column mapping**: Required columns must be mapped for conversion to work
- **Backup**: Original files in input folder remain unchanged

## 🆕 New Features vs JavaScript Version

✅ **GUI Interface** - No command line needed  
✅ **Visual Progress** - See processing status  
✅ **Multi-threading** - Non-blocking UI  
✅ **Better Error Handling** - User-friendly error messages  
✅ **File Management** - Visual file and sheet selection  
✅ **Real-time Logs** - See what's happening  

## 📝 Version History

- **v1.0** - Initial Python GUI version with all JavaScript features
- Support for all unit conversions (K, LBS, etc.)
- Optimized performance for large files
- Enhanced user experience with GUI

---

**Made with ❤️ for GlobalWitz by GitHub Copilot**
