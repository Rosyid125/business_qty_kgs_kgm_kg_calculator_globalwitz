@echo off
setlocal EnableDelayedExpansion

echo ===============================================
echo    Building Business Quantity Converter
echo                 (OneDir Mode)
echo ===============================================
echo.

REM [1/5] Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python not detected! Please install Python and add to PATH
    pause
    exit /b 1
)

REM [2/5] Check and install PyInstaller
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

REM [3/5] Clean previous builds
echo Cleaning previous build folders...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

REM [4/5] Build executable with PyInstaller
echo Building standalone executable...
pyinstaller --onedir --windowed --name="BusinessQuantityConverter" --add-data="README.md;." business_quantity_converter.py

if errorlevel 1 (
    echo ❌ Build failed!
    pause
    exit /b 1
)

REM [5/5] Setup Production_Release folder
echo.
echo Creating Production_Release folder...
if exist Production_Release rmdir /s /q Production_Release
mkdir Production_Release
mkdir Production_Release\input
mkdir Production_Release\output

echo Copying executable and files...
xcopy /e /i /h /y dist\BusinessQuantityConverter Production_Release\BusinessQuantityConverter >nul
if exist README.md copy README.md Production_Release\ >nul

echo Creating user guide...
echo # Business Quantity Converter - User Guide > Production_Release\USER_GUIDE.txt
echo. >> Production_Release\USER_GUIDE.txt
echo HOW TO USE: >> Production_Release\USER_GUIDE.txt
echo 1. Put your Excel files in the 'input' folder >> Production_Release\USER_GUIDE.txt
echo 2. Run BusinessQuantityConverter.exe inside the BusinessQuantityConverter folder >> Production_Release\USER_GUIDE.txt
echo 3. Follow the instructions on screen >> Production_Release\USER_GUIDE.txt
echo 4. Results will appear in 'output' folder >> Production_Release\USER_GUIDE.txt

echo.
echo ===============================================
echo ✅ OneDir Build Complete!
echo ===============================================
echo 📁 Output: Production_Release
echo   - BusinessQuantityConverter (app folder)
echo   - input (place Excel files here)
echo   - output (result will be saved here)
echo   - USER_GUIDE.txt
echo   - README.md
echo.
echo 🚀 Ready for distribution!
echo.
pause

start "" "Production_Release"
