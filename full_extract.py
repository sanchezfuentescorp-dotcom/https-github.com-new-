import pandas as pd, sys

FILE = 'Dashboard Culinary Depurado Santiago.xlsx'
sys.stdout.reconfigure(encoding='utf-8')

# ── HOJA 1: RAMOS (header at row index 1, col 0 = NaN leadeer) ──────────
df1 = pd.read_excel(FILE, sheet_name='DETALLE RAMO MADRE ', header=None)
df1 = df1.iloc[2:, 1:7].reset_index(drop=True)  # skip 2 header rows, skip col 0
df1.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Seccion','N_Alumnos']
df1 = df1.dropna(subset=['Codigo_ID'])
df1['Semestre'] = df1['Semestre'].astype(int)
df1['Seccion'] = df1['Seccion'].astype(int)
df1['N_Alumnos'] = df1['N_Alumnos'].astype(int)
print('=== HOJA 1: RAMOS MADRE ===')
print(df1.to_string(index=False))
print()

# ── HOJA 2: SUB-RAMOS (col 0=NaN leader, col 6=trailing NaN) ────────────
df2 = pd.read_excel(FILE, sheet_name='DETALLE DE SUB-RAMOS MENSUALES', header=None)
df2 = df2.iloc[1:, 1:7].reset_index(drop=True)  # skip header row, cols 1-6
df2.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Mes','OP']
df2['Semestre'] = df2['Semestre'].ffill()
df2['Codigo_ID'] = df2['Codigo_ID'].ffill()
df2['Siglas'] = df2['Siglas'].ffill()
df2['Nombre'] = df2['Nombre'].ffill()
df2['Mes'] = df2['Mes'].ffill()
df2 = df2.dropna(subset=['OP'])
df2['Semestre'] = df2['Semestre'].astype(int)
print('=== HOJA 2: SUB-RAMOS ===')
print(df2.to_string(index=False))
print()

# ── HOJA 3: INGREDIENTES (col 0=NaN leader) ──────────────────────────────
df3 = pd.read_excel(FILE, sheet_name='INGREDIENTES DE SUB-RAMOS', header=None)
df3 = df3.iloc[1:, 1:7].reset_index(drop=True)  # skip header, col 1-6
df3.columns = ['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3 = df3.dropna(subset=['Ingrediente'])
print('=== HOJA 3: INGREDIENTES ===')
print(f'Total: {len(df3)} filas')
print(f'Ramos: {sorted(df3["Ramo"].unique())}')
print(f'Sub-Ramos: {len(df3["Sub_Ramo"].unique())} unicos')
print(f'Ingredientes: {len(df3["Ingrediente"].unique())} unicos')
print(f'Categorias: {sorted(df3["Categoria"].dropna().unique())}')
print()
print(df3.to_string(index=False))
print()

# ── HOJA 4: PRECIOS ──────────────────────────────────────────────────────
df4 = pd.read_excel(FILE, sheet_name='DETALLE PRECIO INGREDIENTES', header=None)
df4 = df4.iloc[1:, 0:8].reset_index(drop=True)
df4.columns = ['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4 = df4.dropna(subset=['Ingrediente'])
print('=== HOJA 4: PRECIOS ===')
print(f'Total: {len(df4)} ingredientes')
print(f'Categorias: {sorted(df4["Categoria"].dropna().unique())}')
print(f'Proveedores: {sorted(df4["Proveedor"].dropna().unique())}')
print()
print(df4.to_string(index=False))
