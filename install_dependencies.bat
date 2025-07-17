@echo off
echo ===============================================
echo    Installing Dependencies
echo ===============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python not detected! Please install Python and add to PATH
    pause
    exit /b 1
)

echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo ✅ Dependencies installed successfully!
echo.
echo You can now run build.bat or build_onefile.bat
pause
