@echo off
echo ===============================================
echo    Building Production Version
echo    Business Quantity to KG Converter
echo ===============================================
echo.

echo [1/4] Installing PyInstaller...
pip install pyinstaller
if errorlevel 1 (
    echo ❌ Failed to install PyInstaller
    echo Please check your internet connection and try again
    pause
    exit /b 1
)

echo.
echo [2/4] Building standalone executable...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name="BusinessQuantityConverter" ^
    --add-data="README.md;." ^
    business_quantity_converter.py

if errorlevel 1 (
    echo ❌ Build failed
    echo Please check if business_quantity_converter.py exists
    pause
    exit /b 1
)

echo.
echo [3/4] Creating production folder structure...
if exist "Production_Release" rmdir /s /q Production_Release
mkdir Production_Release
mkdir Production_Release\input
mkdir Production_Release\output

echo.
echo [4/4] Copying files and creating user guide...
copy dist\BusinessQuantityConverter.exe Production_Release\
copy README.md Production_Release\ 2>nul

echo # Business Quantity Converter - User Guide > Production_Release\USER_GUIDE.txt
echo. >> Production_Release\USER_GUIDE.txt
echo HOW TO USE: >> Production_Release\USER_GUIDE.txt
echo 1. Put your Excel files in the 'input' folder >> Production_Release\USER_GUIDE.txt
echo 2. Double-click BusinessQuantityConverter.exe >> Production_Release\USER_GUIDE.txt
echo 3. Follow the GUI instructions >> Production_Release\USER_GUIDE.txt
echo 4. Converted files will appear in the 'output' folder >> Production_Release\USER_GUIDE.txt
echo. >> Production_Release\USER_GUIDE.txt
echo SYSTEM REQUIREMENTS: >> Production_Release\USER_GUIDE.txt
echo - Windows 7 or later >> Production_Release\USER_GUIDE.txt
echo - No additional software needed >> Production_Release\USER_GUIDE.txt
echo. >> Production_Release\USER_GUIDE.txt
echo SUPPORT: >> Production_Release\USER_GUIDE.txt
echo - Check README.md for detailed instructions >> Production_Release\USER_GUIDE.txt
echo - Report issues to your IT administrator >> Production_Release\USER_GUIDE.txt

echo.
echo ===============================================
echo    Production Build Complete!
echo ===============================================
echo.
echo 📁 Files created in Production_Release folder:
echo ✅ BusinessQuantityConverter.exe - Main application
echo ✅ input folder - Put Excel files here  
echo ✅ output folder - Converted files appear here
echo ✅ USER_GUIDE.txt - Simple instructions
echo ✅ README.md - Detailed documentation
echo.
echo 🚀 Ready for distribution!
echo.
echo You can now:
echo - ZIP the Production_Release folder
echo - Share it with users
echo - Users just need to extract and run the .exe
echo.
pause

echo.
echo Opening Production_Release folder...
start "" "Production_Release"
