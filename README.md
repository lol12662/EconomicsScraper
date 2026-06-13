# WSJ Treasury & Barchart Futures Scraper

A desktop application that scrapes live U.S. Treasury bond data from the Wall Street Journal and financial futures quotes from Barchart, exports them to Excel, and visualizes the data in an interactive chart — all from a simple GUI.

---

## Features

- **Two data sources** selectable from the GUI:
  - **WSJ Treasuries** — scrapes Maturity, Coupon, and Asked Yield from the WSJ bond page and computes a full set of bond analytics
  - **Barchart Futures** — scrapes Symbol, Contract Name, Latest, Change, Volume, and Time for financial futures contracts
- **Bond analytics** computed automatically for WSJ data: PV, Modified & Macaulay Duration, Simulated Return, Effective Duration, and more
- **Treasury price conversion** — Barchart fractional prices (e.g. `97-08`, `109-162`) are automatically converted to decimals
- **Deduplication** — repeated rows in Barchart data are removed automatically
- **Excel export** with a formula reference sheet prepended to WSJ data
- **Interactive chart** (WSJ source) — line graph with selectable series and configurable X axis
- **Real-time log panel** with save-to-file option
- **Resilient scraping** — three-strategy fallback system handles JavaScript-rendered pages and anti-bot protection

---

## Project Files

| File | Description |
|------|-------------|
| `wsj_treasury_scraper_fixed.py` | Core scraping and data processing logic |
| `wsj_treasury_gui_fixed.py` | Desktop GUI (Tkinter + matplotlib) |
| `WSJ_Treasury_GUI.spec` | PyInstaller build spec for packaging |
| `build_mac.sh` | Build script for macOS |
| `build_windows_exe_fixed.bat` | Build script for Windows |
| `debug_mac.sh` | Diagnostic script for macOS troubleshooting |

---

## Requirements

- **Python 3.8+** — download from [python.org](https://www.python.org/downloads/)
- The build scripts install all Python libraries automatically

### Python Libraries

```
pandas
openpyxl
requests
playwright
matplotlib
pyinstaller
```

---

## Quick Start

### Running directly (no build needed)

Make sure all files are in the same folder, then install dependencies once:

```bash
pip install pandas openpyxl requests playwright matplotlib
python -m playwright install chromium
```

Then run the GUI:

```bash
python3 wsj_treasury_gui_fixed.py
```

Or run from the command line without the GUI:

```bash
# WSJ Treasuries
python3 wsj_treasury_scraper_fixed.py --source wsj

# Barchart Futures
python3 wsj_treasury_scraper_fixed.py --source barchart

# Custom output file
python3 wsj_treasury_scraper_fixed.py --source barchart --output ~/Desktop/futures.xlsx

# WSJ with a specific reference date
python3 wsj_treasury_scraper_fixed.py --source wsj --reference-date 2026-06-01
```

---

## Building a Standalone App

### macOS(Work In Progress)

Place all files in the same folder, open Terminal, navigate to that folder, and run:

```bash
bash build_mac.sh
```

This installs all dependencies and builds `dist/WSJ_Treasury_GUI.app`. Double-click the `.app` to launch.

> **First time on macOS:** If you see a security warning, right-click the `.app` → **Open** → **Open**.

### Windows

Double-click `build_windows_exe_fixed.bat`. This installs all dependencies and builds `dist/WSJ_Treasury_GUI.exe`.

> **Requirement:** Python must be installed and on your system PATH before running the `.bat`.

---

## Using the GUI

1. **Select a data source** using the radio buttons at the top (WSJ Treasuries or Barchart Futures)
2. **Set an output file** — defaults to the same folder as the scripts
3. *(WSJ only)* Optionally enter a **reference date** in `YYYY-MM-DD` format to override the page date
4. Click **Run scrape** — progress appears in the log panel on the right
5. When complete, the **Preview tab** shows the data table and *(WSJ only)* the **Chart tab** shows an interactive line graph

### Chart Tab (WSJ only)

- Check/uncheck series on the left panel to show or hide individual metrics
- Change the **X axis** dropdown to plot against Maturity, Article Date, or Last Payment Date
- Use the matplotlib toolbar at the bottom to zoom, pan, or save the chart as an image

### Log Panel

- All scraping steps and diagnostic messages appear here in real time
- Click **Save log…** to save the full log to a `.txt` file
- Click **Clear log** to reset it

---

## WSJ Bond Analytics

The following columns are computed and written to Excel for each Treasury bond:

| Column | Formula |
|--------|---------|
| Coupon Payment | Coupon / 2 × 1000 |
| PV0 | PV(Ask Yield/2, Payments Remaining, Coupon Payment, 1000) |
| PV1 | PV((Ask Yield+1%)/2, Payments Remaining−2, Coupon Payment, 1000) |
| P0 | PV0 × (1 + Ask Yield/2)^(Days Since Last Payment / 182) |
| P1 | PV1 × (1 + (Ask Yield+1%)/2)^(Days Since Last Payment / 182) |
| Simulated Return | (P1 − P0 + Coupon × 10) / P0 |
| Macaulay Duration | DURATION(Today, Maturity, Coupon, Ask Yield, 2) |
| Modified Duration | MDURATION(Today, Maturity, Coupon, Ask Yield, 2) |
| Effective Duration | (PDN − PUP) / (2 × 0.01 × P0) |

---

## Barchart Price Conversion

Treasury futures prices on Barchart use fractional 32nds notation. The scraper converts these automatically:

| Format | Example | Calculation | Decimal |
|--------|---------|-------------|---------|
| `XX-YY` | `97-08` | 97 + 8/32 | 97.25 |
| `XX-YY+` | `97-08+` | 97 + 8/32 + 0.5/32 | 97.265625 |
| `XX-YYZ` | `109-162` | 109 + 16/32 + 2/(8×32) | 109.507813 |

---

## Barchart Futures Contracts

The following front-month contracts are queried:

| Symbol | Contract |
|--------|----------|
| ZT\*0 | 2-Year T-Note |
| ZF\*0 | 5-Year T-Note |
| ZN\*0 | 10-Year T-Note |
| ZB\*0 | 30-Year T-Bond |
| UB\*0 | Ultra T-Bond |
| SR3\*0 | SOFR 3-Month |
| FF\*0 | 30-Day Fed Funds |
| ZQ\*0 | 30-Day Fed Funds (alt) |

---

## Scraping Architecture

The Barchart scraper uses a three-strategy fallback to handle JavaScript-rendered pages:

1. **Proxima internal API** — calls Barchart's internal `/proxima/core/getQuote` endpoint using a session pre-loaded with cookies and XSRF token from the page
2. **Legacy marketdata API** — calls `marketdata.websol.barchart.com/getQuote.json` with Barchart's embedded guest key
3. **Playwright browser interception** — launches a headless Chromium browser, intercepts all XHR/fetch responses, harvests real session cookies, and retries the API calls with them

The WSJ scraper first tries a direct `requests` fetch, then falls back to Playwright if the page requires JavaScript rendering.

---

## Troubleshooting

**App quits immediately on macOS**
Run `bash debug_mac.sh` — this launches the app directly in Terminal so any error prints on screen.

**"Could not extract futures table" error**
Make sure Playwright and Chromium are installed:
```bash
pip install playwright
python -m playwright install chromium
```

**Barchart Change column is empty**
Barchart sometimes omits the change field from their API response. Check the log for `Raw payload keys:` to see what fields were actually returned.

**WSJ page not loading**
The WSJ page occasionally requires a login. Try opening the URL in a browser first to confirm it's accessible, then re-run the scrape.
