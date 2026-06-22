import yfinance as yf
import pandas as pd
from fredapi import Fred

# ── Settings ──────────────────────────────────────────────────────────────────
# Get a free FRED API key at: https://fred.stlouisfed.org/docs/api/api_key.html
FRED_API_KEY = "64ef5699f1878ba7483a60e27db515de "

tickers      = ['HYG', 'SPY', '^TNX']
fred_series  = ['DGS1', 'DTB3', 'DGS10', 'DBAA']
start_date   = "2025-01-02"
end_date     = "2025-06-30"
output_file  = "HighYieldData.csv"
# ──────────────────────────────────────────────────────────────────────────────

# 1. Yahoo Finance
print("Downloading Yahoo Finance data...")
yahoo_data = yf.download(
    tickers,
    start=start_date,
    end=end_date,
    auto_adjust=False,
    progress=False
)['Adj Close']
yahoo_data = yahoo_data.dropna()
print(f"  ✓ Yahoo: {len(yahoo_data)} rows")

# 2. FRED via fredapi (much more reliable than pandas_datareader)
print("Downloading FRED data...")
fred = Fred(api_key=FRED_API_KEY)
fred_data = pd.DataFrame()
for series in fred_series:
    try:
        fred_data[series] = fred.get_series(series, start_date, end_date)
        print(f"  ✓ FRED: {series}")
    except Exception as e:
        print(f"  ✗ FRED: {series} failed — {e}")

# 3. Merge — inner join keeps only dates present in both sources
combined = yahoo_data.join(fred_data, how='inner').dropna()

# 4. Format date column as YYYYMMDD (matches what Bond_Regress expects)
combined = combined.reset_index()
combined.rename(columns={combined.columns[0]: 'Time'}, inplace=True)
combined['Time'] = pd.to_datetime(combined['Time']).dt.strftime('%Y%m%d')

# 5. Save
combined.to_csv(output_file, index=False)
print(f"\nSaved {len(combined)} rows to '{output_file}'")
print(f"Columns: {list(combined.columns)}")
