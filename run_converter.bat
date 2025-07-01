@echo off
echo ===============================================
echo    Business Quantity to KG Converter
echo                (Python GUI)
echo ===============================================
echo.

echo [1/3] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERROR: Python is not installed or not in PATH
    echo.
    echo 💡 Solutions:
    echo    1. Install Python from: https://python.org
    echo    2. Make sure "Add Python to PATH" is checked
    echo    3. Restart this script after installation
    echo.
    pause
    exit /b 1
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo ✅ Found: %%i
)

echo.
echo [2/3] Checking required modules...

echo Checking pandas...
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo ❌ pandas not found
    goto :install_modules
) else (
    echo ✅ pandas available
)

echo Checking openpyxl...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ❌ openpyxl not found
    goto :install_modules
) else (
    echo ✅ openpyxl available
)

echo Checking xlrd...
python -c "import xlrd" >nul 2>&1
if errorlevel 1 (
    echo ❌ xlrd not found
    goto :install_modules
) else (
    echo ✅ xlrd available
)

echo.
echo [3/3] Starting Business Quantity Converter GUI...
echo.
echo 🚀 Launching application...
echo    (GUI window will open in a moment)
echo.

python business_quantity_converter.py

echo.
echo Application closed.
echo Thank you for using Business Quantity Converter!
pause
exit /b 0

:install_modules
echo.
echo ⚠️  Some required modules are missing!
echo.
echo Would you like to install them automatically? (y/n)
set /p choice="Enter your choice: "

if /i "%choice%"=="y" (
    echo.
    echo Installing modules automatically...
    call install_modules.bat
    if errorlevel 1 (
        echo.
        echo ❌ Module installation failed!
        echo Please run install_modules.bat manually
        pause
        exit /b 1
    )
    echo.
    echo ✅ Modules installed! Restarting converter...
    echo.
    goto :start_converter
) else (
    echo.
    echo Please install required modules by running:
    echo    install_modules.bat
    echo.
    echo Or install manually:
    echo    pip install pandas openpyxl xlrd
    pause
    exit /b 1
)

:start_converter
echo [3/3] Starting Business Quantity Converter GUI...
echo.
echo 🚀 Launching application...
python business_quantity_converter.py
echo.
echo Application closed.
pause
