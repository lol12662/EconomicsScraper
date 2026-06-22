import pandas as pd
import statsmodels.formula.api as smf

# ── Settings ──────────────────────────────────────────────────────────────────
input_file = "HighYieldData.csv"
# ──────────────────────────────────────────────────────────────────────────────

# 1. Read and sort data
df = pd.read_csv(input_file, dtype={'Time': str})
df['Date'] = pd.to_datetime(df['Time'], format='%Y%m%d')
df = df.sort_values('Date').reset_index(drop=True)

# 2. Compute returns
df['RHYG'] = df['HYG'] / df['HYG'].shift(1) - 1   # simple return on HYG
df['RSP']  = df['SPY'] / df['SPY'].shift(1) - 1   # simple return on SPY
df['CHI']  = df['TNX'] - df['TNX'].shift(1)        # change in 10yr yield

# Drop first row (NaN from shift)
df_ret = df[['Date', 'RHYG', 'RSP', 'CHI']].dropna().reset_index(drop=True)

# 3. Preview data
print("First 7 rows of return data:")
print(df_ret.head(7).to_string(index=False))

print("\nLast 7 rows of return data:")
print(df_ret.tail(7).to_string(index=False))

# 4. Regression Model 1: RHYG ~ RSP
print("\n--- Model 1: RHYG ~ RSP ---")
model1 = smf.ols('RHYG ~ RSP', data=df_ret).fit()
print(model1.summary())

# 5. Regression Model 2: RHYG ~ RSP + CHI
print("\n--- Model 2: RHYG ~ RSP + CHI ---")
model2 = smf.ols('RHYG ~ RSP + CHI', data=df_ret).fit()
print(model2.summary())
