import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}
FILE='Dashboard Culinary Depurado Santiago.xlsx'
def getall(t,sel='*',extra=''):
    out,off=[],0
    while True:
        r=requests.get(f'{URL}/rest/v1/{t}?select={sel}{extra}',headers={**H,'Range':f'{off}-{off+999}'})
        d=r.json()
        if not isinstance(d,list) or not d: break
        out+=d
        if len(d)<1000: break
        off+=1000
    return out

# ── A. 28 filas cantidad_total = 0 ─────────────────────────────────────
print('[A] INGREDIENTES-SESIÓN CON CANTIDAD 0 (28)')
print('─'*70)
si=getall('sesion_ingrediente','sesion_id,ingrediente_id,cantidad_unit,cantidad_total,ingrediente(nombre)')
cero=[r for r in si if not r['cantidad_total'] or r['cantidad_total']==0]
import collections
nombres=collections.Counter(r['ingrediente']['nombre'] for r in cero)
for n,c in nombres.most_common():
    print(f'  {c:>3}×  {n}')
# Verificar en Excel: esos ingredientes tenían cantidad 0/NaN?
df3=pd.read_excel(FILE,'INGREDIENTES DE SUB-RAMOS',header=None).iloc[1:,1:7].reset_index(drop=True)
df3.columns=['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3['Ingrediente']=df3['Ingrediente'].astype(str).str.strip().str.upper()
df3['Cantidad']=pd.to_numeric(df3['Cantidad'],errors='coerce')
print('\n  Verificación en Excel (cantidad original):')
for n in nombres:
    filas=df3[df3['Ingrediente']==n.upper()]
    cant0=filas[(filas['Cantidad'].isna())|(filas['Cantidad']==0)]
    print(f'    {n[:35]:37} apariciones:{len(filas)} con_cant_0/NaN:{len(cant0)}')

# ── B. Talleres en $0 ──────────────────────────────────────────────────
print('\n[B] TALLERES SIN COSTO – ¿tienen ingredientes?')
print('─'*70)
ses=getall('sesion','id,taller_tipo_id,cod_clase')
talls={t['id']:t['codigo'] for t in getall('taller_tipo','id,codigo')}
si_ses={r['sesion_id'] for r in si}
from collections import defaultdict
tall_ses=defaultdict(int); tall_con_ing=defaultdict(int)
for s in ses:
    cod=talls.get(s['taller_tipo_id'])
    tall_ses[cod]+=1
    if s['id'] in si_ses: tall_con_ing[cod]+=1
for cod in ['TMA','TAS','TIR','TMI','TEAC']:
    print(f'  {cod:6} sesiones:{tall_ses[cod]:>3}  con ingredientes:{tall_con_ing[cod]:>3}')

# Confirmar en Excel: esos talleres tienen ingredientes en hoja3?
ramos_h3=set(df3['Ramo'].astype(str).str.strip().unique())
print(f'\n  Ramos presentes en hoja INGREDIENTES: {sorted(ramos_h3)}')
print(f'  Talleres SIN receta en Excel: {sorted(set([talls[s["taller_tipo_id"]] for s in ses]) - ramos_h3)}')

# ── C. Resumen global de costos ────────────────────────────────────────
print('\n[C] TOTALES GLOBALES')
print('─'*70)
vs=getall('v_costo_sesion','costo_total,num_alumnos')
tot=sum(s['costo_total'] or 0 for s in vs)
alu=sum(s['num_alumnos'] or 0 for s in vs)
print(f'  Costo total materia prima (todas las sesiones): ${tot:,.0f}'.replace(',','.'))
print(f'  Total alumnos-sesión: {alu}')
print(f'  Sesiones con costo > 0: {sum(1 for s in vs if (s["costo_total"] or 0)>0)}/{len(vs)}')
