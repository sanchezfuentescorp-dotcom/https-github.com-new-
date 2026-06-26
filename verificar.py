import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')
FILE = 'Dashboard Culinary Depurado Santiago.xlsx'

df3 = pd.read_excel(FILE, sheet_name='INGREDIENTES DE SUB-RAMOS', header=None)
df3 = df3.iloc[1:, 1:7].reset_index(drop=True)
df3.columns = ['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3 = df3.dropna(subset=['Ingrediente'])
df3['Ingrediente'] = df3['Ingrediente'].str.strip().str.upper()

df4 = pd.read_excel(FILE, sheet_name='DETALLE PRECIO INGREDIENTES', header=None)
df4 = df4.iloc[1:, 0:8].reset_index(drop=True)
df4.columns = ['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4 = df4.dropna(subset=['Ingrediente'])
df4['Ingrediente'] = df4['Ingrediente'].str.strip().str.upper()
df4['Categoria'] = df4['Categoria'].apply(lambda x: None if str(x).startswith('#') else str(x).strip().upper() if pd.notna(x) else None)
df4 = df4[df4['Categoria'].notna()]

ing_h3 = set(df3['Ingrediente'].unique())
ing_h4 = set(df4['Ingrediente'].unique())

solo_h3 = sorted(ing_h3 - ing_h4)  # en sesiones pero sin precio
solo_h4 = sorted(ing_h4 - ing_h3)  # con precio pero sin sesion

print(f"Ingredientes en Hoja3 (sesiones):  {len(ing_h3)}")
print(f"Ingredientes en Hoja4 (precios):   {len(ing_h4)}")
print(f"En ambas hojas:                    {len(ing_h3 & ing_h4)}")
print()
print(f"EN SESIONES PERO SIN PRECIO ({len(solo_h3)}):")
for n in solo_h3:
    print(f"  - {n}")
print()
print(f"CON PRECIO PERO SIN SESION ({len(solo_h4)}) - solo muestra primeros 20:")
for n in solo_h4[:20]:
    print(f"  - {n}")
