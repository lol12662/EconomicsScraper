#!/bin/bash
# Build the WSJ Treasury / Barchart Futures scraper app on macOS.
# Requirements: Python 3.8+ must be installed.
# Run with: bash build_mac.sh

set -e

# ── Check Python is available ──────────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo ""
    echo "ERROR: Python 3 is not installed or not on your PATH."
    echo "Download it from https://www.python.org/downloads/"
    echo ""
    exit 1
fi

echo "Using $(python3 --version)"

# ── Install dependencies ───────────────────────────────────────────────────────
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller pandas openpyxl requests playwright

# ── Install Playwright's Chromium browser ─────────────────────────────────────
python3 -m playwright install chromium

# ── Build the .app bundle ──────────────────────────────────────────────────────
python3 -m PyInstaller --noconfirm --onefile --windowed --name WSJ_Treasury_GUI \
    --add-data "wsj_treasury_scraper_fixed.py:." \
    wsj_treasury_gui_fixed.py

echo ""
echo "Build complete. The app is at dist/WSJ_Treasury_GUI"
echo "You can double-click it or run: open dist/WSJ_Treasury_GUI"
