@echo off
setlocal

REM Build the EXE with Playwright support.
REM Run this on Windows after installing Python.

python -m pip install --upgrade pip
python -m pip install pyinstaller pandas openpyxl requests playwright
python -m playwright install chromium

pyinstaller --noconfirm --onefile --windowed --name WSJ_Treasury_GUI ^
  --add-data "%LOCALAPPDATA%\ms-playwright;ms-playwright" ^
  --hidden-import playwright.sync_api ^
  --add-data "wsj_treasury_scraper_fixed.py;." ^
  wsj_treasury_gui_fixed.py

echo.
echo Build complete. The EXE is in dist\WSJ_Treasury_GUI.exe
pause
