@echo off
echo ===============================================
echo    Python Modules Installer for
echo    Business Quantity to KG Converter
echo ===============================================
echo.

echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: Python is not installed or not in PATH
    echo.
    echo 💡 Please install Python from: https://python.org
    echo    ⚠️  IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo ✅ Found: %%i
)

echo.
echo [2/5] Checking pip (Python package installer)...
pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: pip is not available
    echo 💡 pip should come with Python. Try reinstalling Python.
    pause
    exit /b 1
) else (
    for /f "tokens=*" %%i in ('pip --version 2^>^&1') do echo ✅ Found: %%i
)

echo.
echo [3/5] Upgrading pip to latest version...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo ⚠️  Warning: Could not upgrade pip, continuing anyway...
) else (
    echo ✅ Pip upgraded successfully
)

echo.
echo [4/5] Installing required Python modules...
echo.

echo Installing pandas (Excel data processing)...
pip install pandas>=1.5.0
if errorlevel 1 (
    echo ❌ Failed to install pandas
    goto :error
) else (
    echo ✅ pandas installed successfully
)

echo.
echo Installing openpyxl (Excel file support)...
pip install openpyxl>=3.0.0
if errorlevel 1 (
    echo ❌ Failed to install openpyxl
    goto :error
) else (
    echo ✅ openpyxl installed successfully
)

echo.
echo Installing xlrd (Legacy Excel file support)...
pip install xlrd>=2.0.0
if errorlevel 1 (
    echo ❌ Failed to install xlrd
    goto :error
) else (
    echo ✅ xlrd installed successfully
)

echo.
echo [5/5] Verifying installation...
echo.

echo Checking pandas...
python -c "import pandas; print('✅ pandas version:', pandas.__version__)" 2>nul
if errorlevel 1 (
    echo ❌ pandas verification failed
    goto :error
)

echo Checking openpyxl...
python -c "import openpyxl; print('✅ openpyxl version:', openpyxl.__version__)" 2>nul
if errorlevel 1 (
    echo ❌ openpyxl verification failed
    goto :error
)

echo Checking xlrd...
python -c "import xlrd; print('✅ xlrd version:', xlrd.__version__)" 2>nul
if errorlevel 1 (
    echo ❌ xlrd verification failed
    goto :error
)

echo.
echo ===============================================
echo 🎉 ALL MODULES INSTALLED SUCCESSFULLY!
echo ===============================================
echo.
echo Your system is now ready to run the
echo Business Quantity to KG Converter
echo.
echo Next steps:
echo 1. Put your Excel files in the 'input' folder
echo 2. Double-click 'run_converter.bat' to start the GUI
echo.
echo ===============================================
pause
exit /b 0

:error
echo.
echo ===============================================
echo ❌ INSTALLATION FAILED!
echo ===============================================
echo.
echo Possible solutions:
echo 1. Run this script as Administrator
echo 2. Check your internet connection
echo 3. Make sure Python is properly installed
echo 4. Try manual installation:
echo    pip install pandas openpyxl xlrd
echo.
echo If problems persist, please contact support.
echo ===============================================
pause
exit /b 1
