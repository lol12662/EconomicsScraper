#!/bin/bash
# Double-click this file to install dependencies and build the app.
# First time only: macOS may ask you to allow it in System Preferences > Security.

# Change to the folder where this script lives
cd "$(dirname "$0")"

# ── Check Python is available ──────────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo ""
    echo "ERROR: Python 3 is not installed."
    echo "Download it from https://www.python.org/downloads/ then run this again."
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

echo "Using $(python3 --version)"
echo ""

# ── Install dependencies ───────────────────────────────────────────────────────
echo "Installing Python libraries..."
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller pandas openpyxl requests playwright

# ── Install Playwright's Chromium browser ─────────────────────────────────────
echo ""
echo "Installing Chromium browser for Playwright..."
python3 -m playwright install chromium

# ── Build the app ─────────────────────────────────────────────────────────────
echo ""
echo "Building app..."
python3 -m PyInstaller --noconfirm --onefile --windowed --name WSJ_Treasury_GUI \
    --add-data "wsj_treasury_scraper_fixed.py:." \
    wsj_treasury_gui_fixed.py

echo ""
echo "================================================"
echo "  Build complete!"
echo "  Your app is in the dist/ folder."
echo "  Double-click dist/WSJ_Treasury_GUI to run it."
echo "================================================"
echo ""
read -p "Press Enter to close..."
