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
    echo.
    echo   Your app folder:
    echo   dist\WSJ_Treasury_GUI\
    echo.
    echo   Send the entire dist\WSJ_Treasury_GUI\ folder
    echo   to anyone - they just double-click the .exe
    echo   No Python needed on their machine.
    echo.
    echo   NOTE: First Barchart scrape on a new machine
    echo   will auto-install Chromium ^(~150 MB, once only^).
    echo ================================================
    explorer "dist\WSJ_Treasury_GUI"
) else (
    echo ================================================
    echo   ERROR: EXE not found at expected location:
    echo   dist\WSJ_Treasury_GUI\WSJ_Treasury_GUI.exe
    echo.
    echo   Check the output above for errors.
    echo ================================================
)
pause
