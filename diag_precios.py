import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')
FILE='Dashboard Culinary Depurado Santiago.xlsx'
df4=pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns=['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4['Ingrediente']=df4['Ingrediente'].astype(str).str.strip().str.upper()

busca=['BATON 500 PETIT PAIN','CACAO EN POLVO ALCALINO','LECHE EN POLVO ENTERA',
       'LEVADURA FRESCA','MANTEQUILLA EXTRA SECA','HARINA BIZCOCHERA','ACEITE DE COCO',
       'ACEITUNAS NEGRAS S/HUESO','CLAVO DE OLOR ENTERO','SAL DE MESA LOBOS','MIEL']
print(f"{'INGREDIENTE':32} {'Cant_T':>9} {'Formato':>8} {'Med':>4} {'Neto':>9} {'Neto/Form':>10}")
print('─'*80)
for b in busca:
    rows=df4[df4['Ingrediente'].str.contains(b,na=False,regex=False)]
    for _,r in rows.iterrows():
        nf = r['Neto']/r['Formato'] if pd.notna(r['Formato']) and r['Formato'] not in (0,None) else None
        print(f"{r['Ingrediente'][:32]:32} {str(r['Cant_Total']):>9} {str(r['Formato']):>8} {str(r['Medida']):>4} {str(r['Neto']):>9} {nf if nf is None else round(nf,2):>10}")

print('\n── DISTRIBUCIÓN DE MEDIDAS ──')
print(df4['Medida'].value_counts().to_string())
