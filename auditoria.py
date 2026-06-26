import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, pandas as pd, sys, re, math
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

def cnt(t,q=''):
    r=requests.get(f'{URL}/rest/v1/{t}?{q}',headers={**H,'Prefer':'count=exact','Range':'0-0'})
    return int(r.headers.get('content-range','0/0').split('/')[-1])

OK='✓'; BAD='✗'; WARN='⚠'
def line(s='─',n=70): print(s*n)

print('╔'+'═'*68+'╗')
print('║  AUDITORÍA COMPLETA – CULINARY SANTIAGO (Supabase)'.ljust(68)+'║')
print('╚'+'═'*68+'╝')

# ── 1. CONTEOS ─────────────────────────────────────────────────────────
print('\n[1] CONTEOS DE REGISTROS')
line()
tablas={'categoria':34,'proveedor':22,'taller_tipo':14,'ingrediente':613,'sesion':1135,'sesion_ingrediente':23213}
for t,esp in tablas.items():
    real=cnt(t)
    flag=OK if real==esp else BAD
    print(f'  {flag} {t:22} {real:>6}  (esperado {esp})')

# ── 2. INTEGRIDAD REFERENCIAL ──────────────────────────────────────────
print('\n[2] INTEGRIDAD REFERENCIAL (huérfanos)')
line()
ings=getall('ingrediente','id,nombre,costo_neto,unidad,categoria_id,proveedor_id')
ses=getall('sesion','id,cod_clase,taller_tipo_id,num_alumnos,seccion,fecha')
si=getall('sesion_ingrediente','id,sesion_id,ingrediente_id,cantidad_unit,cantidad_total,unidad')
cats={c['id'] for c in getall('categoria','id')}
provs={p['id'] for p in getall('proveedor','id')}
talls={t['id'] for t in getall('taller_tipo','id')}
ing_ids={i['id'] for i in ings}
ses_ids={s['id'] for s in ses}

ing_sin_cat=[i for i in ings if i['categoria_id'] not in cats]
ing_sin_prov=[i for i in ings if i['proveedor_id'] is not None and i['proveedor_id'] not in provs]
ses_sin_tall=[s for s in ses if s['taller_tipo_id'] not in talls]
si_sin_ses=[r for r in si if r['sesion_id'] not in ses_ids]
si_sin_ing=[r for r in si if r['ingrediente_id'] not in ing_ids]
def rep(lbl,lst):
    print(f'  {OK if not lst else BAD} {lbl:42} {len(lst)}')
rep('ingredientes con categoría inexistente',ing_sin_cat)
rep('ingredientes con proveedor inexistente',ing_sin_prov)
rep('sesiones con taller inexistente',ses_sin_tall)
rep('sesion_ingrediente con sesión inexistente',si_sin_ses)
rep('sesion_ingrediente con ingrediente inexistente',si_sin_ing)

# ── 3. VALORES NULOS / ANÓMALOS ────────────────────────────────────────
print('\n[3] VALORES NULOS Y ANÓMALOS')
line()
ing_precio0=[i for i in ings if not i['costo_neto'] or i['costo_neto']==0]
ing_sin_prov_null=[i for i in ings if i['proveedor_id'] is None]
ses_sin_alum=[s for s in ses if not s['num_alumnos'] or s['num_alumnos']==0]
ses_sin_fecha=[s for s in ses if not s['fecha']]
si_cant0=[r for r in si if not r['cantidad_total'] or r['cantidad_total']==0]
print(f'  {WARN if ing_precio0 else OK} ingredientes con precio $0          {len(ing_precio0)}')
print(f'  {WARN if ing_sin_prov_null else OK} ingredientes sin proveedor          {len(ing_sin_prov_null)}')
print(f'  {BAD if ses_sin_alum else OK} sesiones sin alumnos               {len(ses_sin_alum)}')
print(f'  {BAD if ses_sin_fecha else OK} sesiones sin fecha                 {len(ses_sin_fecha)}')
print(f'  {WARN if si_cant0 else OK} ingredientes-sesión con cantidad 0   {len(si_cant0)}')

# ── 4. DUPLICADOS ──────────────────────────────────────────────────────
print('\n[4] DUPLICADOS')
line()
import collections
dup_ing=[k for k,v in collections.Counter(i['nombre'].upper() for i in ings).items() if v>1]
dup_cod=[k for k,v in collections.Counter(s['cod_clase'] for s in ses).items() if v>1]
dup_cat=[k for k,v in collections.Counter(c for c in [x['nombre'] for x in getall('categoria','nombre')]).items() if v>1]
dup_prov=[k for k,v in collections.Counter(p['nombre'].lower() for p in getall('proveedor','nombre')).items() if v>1]
dup_si=[k for k,v in collections.Counter((r['sesion_id'],r['ingrediente_id']) for r in si).items() if v>1]
print(f'  {BAD if dup_ing else OK} ingredientes duplicados (nombre)    {len(dup_ing)} {dup_ing[:3]}')
print(f'  {BAD if dup_cod else OK} cod_clase duplicados                {len(dup_cod)} {dup_cod[:3]}')
print(f'  {BAD if dup_cat else OK} categorías duplicadas               {len(dup_cat)} {dup_cat[:3]}')
print(f'  {WARN if dup_prov else OK} proveedores duplicados (lower)      {len(dup_prov)} {dup_prov[:3]}')
print(f'  {BAD if dup_si else OK} (sesión,ingrediente) duplicados     {len(dup_si)}')

# ── 5. CONSISTENCIA CON EXCEL ──────────────────────────────────────────
print('\n[5] CONSISTENCIA CON EL EXCEL')
line()
df1=pd.read_excel(FILE,'DETALLE RAMO MADRE ',header=None).iloc[2:,1:7].reset_index(drop=True)
df1.columns=['Semestre','Codigo_ID','Siglas','Nombre','Seccion','N_Alumnos']; df1=df1.dropna(subset=['Codigo_ID'])
df2=pd.read_excel(FILE,'DETALLE DE SUB-RAMOS MENSUALES',header=None).iloc[1:,1:7].reset_index(drop=True)
df2.columns=['Semestre','Codigo_ID','Siglas','Nombre','Mes','OP']
for c in df2.columns: df2[c]=df2[c].ffill()
df2=df2.dropna(subset=['OP'])
df4=pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns=['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4=df4.dropna(subset=['Ingrediente']); df4['Ingrediente']=df4['Ingrediente'].str.strip().str.upper()
df4u=df4.drop_duplicates('Ingrediente')

excel_talleres=set(df1['Siglas'].str.strip().unique())
db_talleres={t['codigo'] for t in getall('taller_tipo','codigo')}
print(f'  Talleres Excel:{len(excel_talleres)}  DB:{len(db_talleres)}  faltan:{excel_talleres-db_talleres or "ninguno"}')

# nº sesiones esperado = sum(OPs del taller × secciones del taller)
sec_por_tall=df1.groupby(df1['Siglas'].str.strip()).size().to_dict()  # nº secciones (filas) por sigla
ops_por_tall=df2.groupby(df2['Siglas'].str.strip()).size().to_dict()
esp_ses=sum(ops_por_tall.get(s,0)*sec_por_tall.get(s,0) for s in excel_talleres)
print(f'  Sesiones esperadas (OP×sec): {esp_ses}   DB: {len(ses)}   {OK if esp_ses==len(ses) else BAD}')

# ingredientes Excel con precio
excel_ings=set(df4u['Ingrediente'])
db_ings={i['nombre'].upper() for i in ings}
print(f'  Ingredientes precio Excel:{len(excel_ings)}  en DB:{len(excel_ings & db_ings)}  faltan:{len(excel_ings-db_ings)}')

# ── 6. VERIFICACIÓN DE CÁLCULO (muestra) ───────────────────────────────
print('\n[6] VERIFICACIÓN DE CÁLCULO cantidad_total = cantidad_unit × alumnos')
line()
ses_alum={s['id']:s['num_alumnos'] for s in ses}
errores_calc=0
for r in si[:5000]:
    n=ses_alum.get(r['sesion_id'],0)
    esp=round((r['cantidad_unit'] or 0)*n,3)
    real=round(r['cantidad_total'] or 0,3)
    if abs(esp-real)>0.01: errores_calc+=1
print(f'  {OK if errores_calc==0 else BAD} muestreo 5000 filas: {errores_calc} discrepancias cantidad_total')

# ── 7. RANGOS DE COSTO POR TALLER ──────────────────────────────────────
print('\n[7] COSTO POR ALUMNO – RANGO POR TALLER')
line()
vt=getall('v_costo_taller','codigo,nombre,costo_promedio_alumno,num_sesiones')
def clp(v): return f"${v:,.0f}".replace(',','.')
altos=0
for t in sorted(vt,key=lambda x:-(x['costo_promedio_alumno'] or 0)):
    c=t['costo_promedio_alumno'] or 0
    flag = WARN if c>10000 else (OK if c>0 else WARN)
    if c>10000: altos+=1
    print(f'  {flag} {t["codigo"]:7} {clp(c):>12}/alu  ({t["num_sesiones"]} ses)')
print(f'\n  {WARN if altos else OK} talleres con costo/alumno > $10.000 (sospechoso): {altos}')

# ── 8. VISTAS ──────────────────────────────────────────────────────────
print('\n[8] ESTADO DE VISTAS')
line()
for v in ['v_costo_sesion','v_costo_taller','v_costo_categoria','v_top_ingredientes','v_lista_compras','v_costo_modulo','v_costo_carrera']:
    r=requests.get(f'{URL}/rest/v1/{v}?select=*',headers={**H,'Range':'0-4999'})
    d=r.json(); ok=isinstance(d,list)
    print(f'  {OK if ok else BAD} {v:22} {"OK" if ok else "OVERFLOW/ERROR"}')

print('\n'+'═'*70)
print('FIN AUDITORÍA')
