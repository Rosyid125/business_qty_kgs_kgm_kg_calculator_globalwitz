@echo off
setlocal EnableDelayedExpansion

echo ===============================================
echo    Building Business Quantity Converter
echo              (OneFile Mode)
echo ===============================================
echo.

REM [1/4] Install PyInstaller if not present
echo Checking PyInstaller...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ❌ Failed to install PyInstaller!
        pause
        exit /b 1
    )
)

REM [2/4] Clean old builds
echo Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

REM [3/4] Build with PyInstaller
echo Building standalone executable (OneFile)...
pyinstaller --onefile --windowed --name="BusinessQuantityConverter" --add-data="README.md;." business_quantity_converter.py

if errorlevel 1 (
    echo ❌ Build failed!
    echo Please check if business_quantity_converter.py exists.
    pause
    exit /b 1
)

REM [4/4] Setup Production_Release folder
echo Creating Production_Release folder...
if exist Production_Release rmdir /s /q Production_Release
mkdir Production_Release
mkdir Production_Release\input
mkdir Production_Release\output

echo Copying executable and files...
copy dist\BusinessQuantityConverter.exe Production_Release\ >nul
if exist README.md copy README.md Production_Release\ >nul

echo Creating USER_GUIDE.txt...
(
    echo # Business Quantity Converter - User Guide
    echo.
    echo HOW TO USE:
    echo 1. Put your Excel files in the 'input' folder
    echo 2. Double-click BusinessQuantityConverter.exe
    echo 3. Follow the instructions on screen
    echo 4. Converted files will appear in 'output' folder
    echo.
    echo SYSTEM REQUIREMENTS:
    echo - Windows 7 or later
    echo - No need to install Python or any dependencies
    echo.
    echo SUPPORT:
    echo - See README.md for details
    echo - Contact IT administrator if you experience issues
) > Production_Release\USER_GUIDE.txt

echo.
echo ===============================================
echo ✅ OneFile Build Complete!
echo ===============================================
echo.
echo 📁 Output: Production_Release
echo   - BusinessQuantityConverter.exe
echo   - input\
echo   - output\
echo   - USER_GUIDE.txt
echo   - README.md
echo.
echo 🚀 Ready to distribute!
echo.

start "" "Production_Release"
pause
