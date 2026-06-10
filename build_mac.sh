#!/bin/bash
# Run with: bash build_mac.sh

# Always run from the folder this script lives in
cd "$(dirname "$0")"

# ── Check Python ───────────────────────────────────────────────────────────────
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

# ── Install libraries ──────────────────────────────────────────────────────────
echo "Installing Python libraries..."
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller pandas openpyxl requests playwright

echo ""
echo "Installing Chromium for Playwright..."
python3 -m playwright install chromium

# ── Build using the .spec file ─────────────────────────────────────────────────
echo ""
echo "Building app..."
python3 -m PyInstaller --noconfirm WSJ_Treasury_GUI.spec

# ── Check it worked ────────────────────────────────────────────────────────────
if [ -d "dist/WSJ_Treasury_GUI.app" ]; then
    echo ""
    echo "================================================"
    echo "  Build complete!"
    echo "  App is in the dist/ folder."
    echo "  Opening it now..."
    echo "================================================"
    open dist/WSJ_Treasury_GUI.app
else
    echo ""
    echo "================================================"
    echo "  ERROR: Build failed — app not found in dist/"
    echo "  Check the output above for error messages."
    echo "================================================"
fi

echo ""
read -p "Press Enter to close..."
