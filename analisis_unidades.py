import pandas as pd, sys, math
sys.stdout.reconfigure(encoding='utf-8')
FILE='Dashboard Culinary Depurado Santiago.xlsx'
df4=pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns=['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4['Ingrediente']=df4['Ingrediente'].astype(str).str.strip().str.upper()
df4['Neto']=pd.to_numeric(df4['Neto'],errors='coerce')
df4['Formato']=pd.to_numeric(df4['Formato'],errors='coerce')
df4=df4.dropna(subset=['Ingrediente'])

def medida_norm(m):
    if m is None or (isinstance(m,float) and pd.isna(m)): return 'UN'
    return str(m).strip().upper()

def convertir(neto, formato, medida):
    """Devuelve (costo_por_unidad_base, unidad_base)."""
    m = medida_norm(medida)
    if formato is None or pd.isna(formato) or formato==0: formato=1
    base = neto/formato  # precio por 1 unidad de medida
    # peso
    if m in ('KG',): return base, 'kg'
    if m in ('GR','G','GRS'):
        # Formato==1 + GR con Neto grande = error de etiqueta (es kg)
        if formato==1 and neto>=500: return base, 'kg'
        return base*1000, 'kg'
    # volumen
    if m in ('L','LT'): return base, 'L'
    if m in ('CC','ML'):
        if formato==1 and neto>=500: return base, 'L'
        return base*1000, 'L'
    # conteo
    return base, 'UN'

rows=[]
for _,r in df4.iterrows():
    c,u = convertir(r['Neto'],r['Formato'],r['Medida'])
    rows.append((r['Ingrediente'],r['Neto'],r['Formato'],medida_norm(r['Medida']),round(c,2),u))

# Outliers por unidad
print('═══ PRECIOS ANÓMALOS (posibles errores de dato) ═══\n')
print('  kg > $60.000  |  L > $60.000  |  UN > $5.000\n')
print(f"{'INGREDIENTE':36} {'Neto':>9} {'Form':>6} {'Med':>4} {'$/base':>12} {'base':>4}")
print('─'*78)
umbral={'kg':60000,'L':60000,'UN':5000}
flagged=[]
for (nom,neto,form,med,c,u) in sorted(rows,key=lambda x:-x[4]):
    if c > umbral.get(u,1e9):
        flagged.append((nom,neto,form,med,c,u))
        print(f"{nom[:36]:36} {neto:>9.0f} {str(form):>6} {med:>4} {c:>12,.0f} {u:>4}")
print(f"\nTotal anómalos: {len(flagged)}")

print('\n═══ DISTRIBUCIÓN UNIDAD BASE RESULTANTE ═══')
from collections import Counter
print(Counter(u for *_,u in rows))
