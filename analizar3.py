import pandas as pd

FILE = 'Dashboard Culinary Depurado Santiago.xlsx'

# ── HOJA 2 raw ─────────────────────────────────────────────────────────
df2_raw = pd.read_excel(FILE, sheet_name='DETALLE DE SUB-RAMOS MENSUALES', header=None)
df2_raw = df2_raw.dropna(axis=1, how='all')
print('Hoja2 shape:', df2_raw.shape)
print(df2_raw.head(20).to_string())
print()

# ── HOJA 3 raw ─────────────────────────────────────────────────────────
df3_raw = pd.read_excel(FILE, sheet_name='INGREDIENTES DE SUB-RAMOS', header=None)
df3_raw = df3_raw.dropna(axis=1, how='all')
print('Hoja3 shape:', df3_raw.shape)
print(df3_raw.head(5).to_string())
print()

# ── HOJA 4 raw ─────────────────────────────────────────────────────────
df4_raw = pd.read_excel(FILE, sheet_name='DETALLE PRECIO INGREDIENTES', header=None)
df4_raw = df4_raw.dropna(axis=1, how='all')
print('Hoja4 shape:', df4_raw.shape)
print(df4_raw.head(5).to_string())
