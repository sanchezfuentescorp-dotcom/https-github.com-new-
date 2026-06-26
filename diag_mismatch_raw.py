import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')
FILE='Dashboard Culinary Depurado Santiago.xlsx'

# Hoja 4 cruda
df4=pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns=['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4['Ingrediente']=df4['Ingrediente'].astype(str).str.strip().str.upper()

# Hoja 3 cruda - ver cómo se usan
df3=pd.read_excel(FILE,'INGREDIENTES DE SUB-RAMOS',header=None).iloc[1:,1:7].reset_index(drop=True)
df3.columns=['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3['Ingrediente']=df3['Ingrediente'].astype(str).str.strip().str.upper()

busca=['COLAPEZ','ALGA NORI','PAPEL DE ARROZ','CHALOTA','PAN MOLDE BLANCO LAMINA',
       'GALLETAS DE CHAMPAÑA','VAINA VAINILLA','MERLUZA FRESCA ENTERA','PAN MARRAQUETA','VIENESA']
for b in busca:
    print(f'\n═══ {b} ═══')
    h4=df4[df4['Ingrediente'].str.contains(b,na=False,regex=False)]
    for _,r in h4.iterrows():
        print(f"  HOJA4: Form={r['Formato']} Med={r['Medida']} Neto={r['Neto']} | CantTotal={r['Cant_Total']}")
    h3=df3[df3['Ingrediente'].str.contains(b,na=False,regex=False)]
    ums=h3['UM'].value_counts().to_dict()
    cants=h3['Cantidad'].describe()[['min','max','mean']].to_dict() if len(h3) else {}
    print(f"  HOJA3: usos={len(h3)} UMs={ums} cant(min/max/avg)={ {k:round(v,3) for k,v in cants.items()} }")
