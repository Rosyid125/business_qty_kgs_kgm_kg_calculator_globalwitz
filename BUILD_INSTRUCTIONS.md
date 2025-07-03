# Business Quantity to KG Converter - Production Build

## Quick Start for Developers

### Building Production Version
1. Run `build_production.bat` to create standalone executable
2. Run `create_distribution.bat` to create shareable ZIP file

### Distribution
- Share the `BusinessQuantityConverter_Distribution.zip` file
- Users extract and run `BusinessQuantityConverter.exe`
- No Python or dependencies installation required

## Production Files Structure

```
Production_Release/
├── BusinessQuantityConverter.exe    # Main executable (self-contained)
├── input/                          # Excel files go here
├── output/                         # Converted files appear here
├── USER_GUIDE.txt                  # Simple user instructions
└── README.md                       # Detailed documentation
```

## User Instructions

### For End Users:
1. Extract the ZIP file to any folder
2. Put Excel files in the `input` folder
3. Double-click `BusinessQuantityConverter.exe`
4. Follow the GUI instructions
5. Find converted files in the `output` folder

### System Requirements:
- Windows 7 or later
- No additional software required
- 50MB disk space

## Build Process Details

The production build uses PyInstaller to create a standalone executable that includes:
- Python interpreter
- All required libraries (pandas, openpyxl, xlrd)
- GUI framework (tkinter)
- Application code

This ensures the application runs on any Windows machine without dependencies.

## Troubleshooting

If build fails:
1. Ensure `business_quantity_converter.py` exists
2. Check internet connection (for PyInstaller download)
3. Run as administrator if needed
4. Verify Python and pip are working

## Distribution Tips

- The executable is about 30-50MB (normal for Python apps)
- Users can run it from any location
- No installation process needed
- Works offline after extraction
