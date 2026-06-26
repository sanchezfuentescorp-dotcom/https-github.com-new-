import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, pandas as pd, sys, re, collections
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
def clp(v): return ('$'+f'{v:,.0f}').replace(',','.')

# ── 1. TOTAL GASTO (exacto como dashboard) ─────────────────────────────
print('[1] TOTAL GASTO MATERIA PRIMA (suma costo_total de v_costo_sesion)')
print('─'*64)
vs=getall('v_costo_sesion','costo_total,fecha')
tot=sum(s['costo_total'] or 0 for s in vs)
print(f'   Total actual en BD: {clp(tot)}')
print(f'   Captura usuario:    $143.444.980')
print(f'   Diferencia:         {clp(143444980-tot)}')

# gasto por mes (para entender el 992%)
print('\n   Gasto por mes (fecha):')
porMes=collections.defaultdict(float)
for s in vs: porMes[(s['fecha'] or '')[:7]]+=s['costo_total'] or 0
for m in sorted(porMes):
    print(f'     {m}: {clp(porMes[m])}')

# ── 2. VERIFICACIÓN cantidades Excel vs BD ─────────────────────────────
print('\n[2] FIDELIDAD CANTIDADES: Excel (hoja INGREDIENTES) vs BD')
print('─'*64)
df3=pd.read_excel(FILE,'INGREDIENTES DE SUB-RAMOS',header=None).iloc[1:,1:7].reset_index(drop=True)
df3.columns=['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3=df3.dropna(subset=['Ingrediente'])
df3['Ingrediente']=df3['Ingrediente'].str.strip().str.upper()
df3['Sub_Ramo']=df3['Sub_Ramo'].astype(str).str.strip()
df3['Cantidad']=pd.to_numeric(df3['Cantidad'],errors='coerce').fillna(0)
# consolidar Excel por (sub_ramo, ingrediente) sumando
ING_MAP={'CHUCRUT ENCURTIDO 680 GR (FRASCO)':'CHUCRUT ENCURTIDO','LECHE FRESCA REBECA':'LECHE FRESCA',
 'MALVAVISCOS 240 GR':'MALVAVISCOS','SAL DE CURA #1':'SAL DE CURA',
 'VASOS ACRILICOS DE DEGUSTACION 2 OZ C/TAPA':'VASOS ACRILICOS DE DEGUSTACION 2 OZ',
 'VASOS DE DEGUSTACION 3 OZ C/TAPA':'VASOS DE DEGUSTACION 3 OZ','PAPA':'PAPAS'}
excel_op_ing=collections.defaultdict(float)
for _,r in df3.iterrows():
    op=re.sub(r'\s+','',r['Sub_Ramo']); ing=ING_MAP.get(r['Ingrediente'],r['Ingrediente'])
    excel_op_ing[(op,ing)]+=r['Cantidad']

# BD: cantidad_unit por (op_base, ingrediente). Reconstruir op_base desde cod_clase
ses={s['id']:s for s in getall('sesion','id,cod_clase')}
ings={i['id']:i['nombre'].upper() for i in getall('ingrediente','id,nombre')}
si=getall('sesion_ingrediente','sesion_id,ingrediente_id,cantidad_unit')
# op_base = cod_clase sin el sufijo -MES-Sn
def op_de_cod(cod):
    return re.sub(r'-[A-Z]{3}-S\d+(-\d+)?$','',cod)
bd_op_ing={}
for r in si:
    s=ses.get(r['sesion_id']);
    if not s: continue
    op=op_de_cod(s['cod_clase']); ing=ings.get(r['ingrediente_id'])
    bd_op_ing[(op,ing)]=r['cantidad_unit']  # cantidad_unit es por alumno (igual en todas secciones)

# comparar
difs=0; comparados=0; faltan_bd=0
ejemplos=[]
for (op,ing),cant_x in excel_op_ing.items():
    key=(op,ing)
    if key in bd_op_ing:
        comparados+=1
        if abs((bd_op_ing[key] or 0)-cant_x)>0.001:
            difs+=1
            if len(ejemplos)<10: ejemplos.append((op,ing,cant_x,bd_op_ing[key]))
    else:
        faltan_bd+=1
print(f'   Pares (sub-ramo, ingrediente) en Excel: {len(excel_op_ing)}')
print(f'   Comparados con BD: {comparados}  |  coinciden: {comparados-difs}  |  difieren: {difs}')
print(f'   En Excel pero no en BD: {faltan_bd}')
if ejemplos:
    print('   Ejemplos de diferencia (op, ing, excel, bd):')
    for e in ejemplos: print(f'     {e[0]:14} {e[1][:26]:28} excel={e[2]:.4f} bd={e[3]}')

# ── 3. VERIFICACIÓN precios Excel vs BD ────────────────────────────────
print('\n[3] FIDELIDAD PRECIOS: Excel (sin conversión) vs BD (con conversión/factores)')
print('─'*64)
df4=pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns=['Cat','Ing','Prov','CantT','Form','Med','Neto','Bruto']
df4['Ing']=df4['Ing'].str.strip().str.upper()
df4['Neto']=pd.to_numeric(df4['Neto'],errors='coerce')
df4u=df4.drop_duplicates('Ing')
dbp={i['nombre'].upper():i for i in getall('ingrediente','nombre,costo_neto,unidad')}
en_excel=set(df4u['Ing']); en_bd=set(dbp)
print(f'   Ingredientes con precio en Excel: {len(en_excel)}')
print(f'   Presentes en BD: {len(en_excel & en_bd)}  |  faltan en BD: {len(en_excel-en_bd)}')
if en_excel-en_bd: print(f'   Faltantes: {sorted(en_excel-en_bd)[:10]}')
