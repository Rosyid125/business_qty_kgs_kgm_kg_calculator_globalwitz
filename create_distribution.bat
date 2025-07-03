@echo off
echo ===============================================
echo    Quick Distribution Package Creator
echo ===============================================
echo.

echo Creating distribution package...
if exist "BusinessQuantityConverter_Distribution.zip" del "BusinessQuantityConverter_Distribution.zip"

echo Compressing Production_Release folder...
powershell -command "Compress-Archive -Path 'Production_Release\*' -DestinationPath 'BusinessQuantityConverter_Distribution.zip' -Force"

if exist "BusinessQuantityConverter_Distribution.zip" (
    echo.
    echo ✅ Distribution package created!
    echo 📦 File: BusinessQuantityConverter_Distribution.zip
    echo.
    echo This ZIP file contains everything users need:
    echo - No Python installation required
    echo - No additional dependencies
    echo - Just extract and run
    echo.
    echo Ready to share with users!
    pause
    
    echo Opening file location...
    explorer /select,"BusinessQuantityConverter_Distribution.zip"
) else (
    echo ❌ Failed to create distribution package
    pause
)
