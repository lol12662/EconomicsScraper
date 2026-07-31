@echo off
setlocal
cd /d "%~dp0"

echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas openpyxl requests playwright matplotlib yfinance fredapi statsmodels

echo.
echo Installing Chromium browser...
python -m playwright install chromium

echo.
echo Cleaning previous build...
if exist "build\WSJ_Treasury_GUI" rmdir /s /q "build\WSJ_Treasury_GUI"
if exist "dist\WSJ_Treasury_GUI"  rmdir /s /q "dist\WSJ_Treasury_GUI"

echo.
echo Building EXE...
python -m PyInstaller --noconfirm WSJ_Treasury_GUI.spec

echo.
if exist "dist\WSJ_Treasury_GUI\WSJ_Treasury_GUI.exe" (
    echo ================================================
    echo   Build complete!
    echo   Your app folder:  dist\WSJ_Treasury_GUI\
    echo   Send the whole folder - just double-click the .exe
    echo   No Python needed on the target machine.
    echo ================================================
    explorer "dist\WSJ_Treasury_GUI"
) else (
    echo ================================================
    echo   ERROR: EXE not found.
    echo   Check the output above for errors.
    echo ================================================
)
pause
