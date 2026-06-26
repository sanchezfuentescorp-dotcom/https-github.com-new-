import pandas as pd

FILE = 'Dashboard Culinary Depurado Santiago.xlsx'

# ── HOJA 1: DETALLE RAMO MADRE ─────────────────────────────────────────
df1 = pd.read_excel(FILE, sheet_name='DETALLE RAMO MADRE ', header=1)
df1 = df1.dropna(axis=1, how='all').dropna(subset=[df1.columns[1]])
df1.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Seccion','N_Alumnos']
print('=== HOJA 1: DETALLE RAMO MADRE ===')
print(f'Total filas: {len(df1)}')
print(df1.to_string(index=False))
print()

# ── HOJA 2: DETALLE DE SUB-RAMOS MENSUALES ─────────────────────────────
df2_raw = pd.read_excel(FILE, sheet_name='DETALLE DE SUB-RAMOS MENSUALES', header=None)
# Forward fill the course info columns
df2_raw = df2_raw.dropna(axis=1, how='all')
df2_raw.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Mes','OP']
df2_raw = df2_raw.iloc[1:].reset_index(drop=True)
df2_raw['Semestre'] = df2_raw['Semestre'].ffill()
df2_raw['Codigo_ID'] = df2_raw['Codigo_ID'].ffill()
df2_raw['Siglas'] = df2_raw['Siglas'].ffill()
df2_raw['Nombre'] = df2_raw['Nombre'].ffill()
df2_raw['Mes'] = df2_raw['Mes'].ffill()
df2_raw = df2_raw.dropna(subset=['OP'])
print('=== HOJA 2: DETALLE DE SUB-RAMOS MENSUALES ===')
print(f'Total sesiones: {len(df2_raw)}')
print(df2_raw.to_string(index=False))
print()

# ── HOJA 3: INGREDIENTES DE SUB-RAMOS ──────────────────────────────────
df3 = pd.read_excel(FILE, sheet_name='INGREDIENTES DE SUB-RAMOS', header=0)
df3 = df3.dropna(axis=1, how='all').iloc[:, 1:]
df3.columns = ['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3 = df3.dropna(subset=['Ingrediente'])
print('=== HOJA 3: INGREDIENTES DE SUB-RAMOS ===')
print(f'Total filas: {len(df3)}')
print(f'Ramos unicos: {sorted(df3["Ramo"].dropna().unique())}')
print(f'Sub-Ramos unicos: {len(df3["Sub_Ramo"].dropna().unique())}')
print(f'Ingredientes unicos: {len(df3["Ingrediente"].dropna().unique())}')
print(f'Categorias: {sorted(df3["Categoria"].dropna().unique())}')
print(df3.to_string(index=False))
print()

# ── HOJA 4: DETALLE PRECIO INGREDIENTES ────────────────────────────────
df4 = pd.read_excel(FILE, sheet_name='DETALLE PRECIO INGREDIENTES', header=0)
df4 = df4.dropna(axis=1, how='all')
df4.columns = ['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4 = df4.dropna(subset=['Ingrediente'])
print('=== HOJA 4: DETALLE PRECIO INGREDIENTES ===')
print(f'Total ingredientes: {len(df4)}')
print(f'Categorias: {sorted(df4["Categoria"].dropna().unique())}')
print(f'Proveedores unicos: {sorted(df4["Proveedor"].dropna().unique())}')
print(df4.to_string(index=False))
