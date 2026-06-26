import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, pandas as pd, sys, re, math, time, json
sys.stdout.reconfigure(encoding='utf-8')

URL = os.getenv('SUPABASE_URL')
KEY = os.getenv('SUPABASE_KEY')
REST = f'{URL}/rest/v1'
H = {'apikey': KEY, 'Authorization': f'Bearer {KEY}', 'Content-Type': 'application/json'}
FILE = 'Dashboard Culinary Depurado Santiago.xlsx'

def hdr(prefer=None):
    h = dict(H)
    if prefer: h['Prefer'] = prefer
    return h

def post(tabla, rows, ret=False, on_conflict=None, batch=1000):
    """Inserta rows en lotes. Si ret=True devuelve lista de registros insertados."""
    out = []
    pref = 'return=representation' if ret else 'return=minimal'
    qp = f'?on_conflict={on_conflict}' if on_conflict else ''
    if on_conflict: pref += ',resolution=merge-duplicates'
    for i in range(0, len(rows), batch):
        chunk = rows[i:i+batch]
        r = requests.post(f'{REST}/{tabla}{qp}', headers=hdr(pref), data=json.dumps(chunk))
        if r.status_code not in (200,201,204):
            print(f'  ✗ ERROR POST {tabla} [{i}:{i+len(chunk)}] {r.status_code}: {r.text[:300]}')
            sys.exit(1)
        if ret and r.text:
            out.extend(r.json())
    return out

def delete_all(tabla, key='id'):
    r = requests.delete(f'{REST}/{tabla}?{key}=gt.0', headers=hdr('return=minimal'))
    ok = r.status_code in (200,204)
    print(f'  {"✓" if ok else "✗"} DELETE {tabla} → {r.status_code}' + ('' if ok else f' {r.text[:200]}'))
    return ok

def get_all(tabla, select='*'):
    out, off, page = [], 0, 1000
    while True:
        r = requests.get(f'{REST}/{tabla}?select={select}',
                         headers={**H, 'Range-Unit':'items','Range':f'{off}-{off+page-1}'})
        if r.status_code not in (200,206): break
        data = r.json()
        out.extend(data)
        if len(data) < page: break
        off += page
    return out

# ══════════════════════════════════════════════════════════════════════════
# PARSEO EXCEL (misma lógica que generar_sql.py)
# ══════════════════════════════════════════════════════════════════════════
CAT_COLORS = {
    'ABARROTE':'#f59e0b','AGUA':'#0ea5e9','ALGAS':'#22c55e','ARTICULO COCINA':'#6b7280',
    'BROTES':'#86efac','CABRITO':'#b45309','CERDO':'#f43f5e','CERVEZA':'#ca8a04',
    'CHOCOLATE':'#78350f','COLORANTES':'#ef4444','CONGELADO':'#38bdf8','CONSERVA':'#d97706',
    'EMBUTIDO':'#dc2626','ENCURTIDO':'#f97316','ENDULZANTE':'#fde68a','ESPORAS':'#7c3aed',
    'FRUTA VERDURA':'#10b981','FRUTOS SECOS':'#ea580c','HUEVO':'#fbbf24','JAMON':'#e11d48',
    'JUGO':'#fb923c','LACTEO':'#c4b5fd','LEGUMBRE':'#92400e','LICOR':'#8b5cf6',
    'MASA':'#fcd34d','NITROGENO':'#7dd3fc','PAN':'#fde68a','PATO':'#b45309',
    'PESCADO MARISCO':'#14b8a6','POLLO':'#fef08a','TE':'#86efac','VACUNO':'#7f1d1d',
    'VINO':'#6d28d9','YOGURT':'#f1f5f9',
}
TALL_COLORS = {
    'TIC':'#ef4444','TIP':'#a855f7','TCCM':'#3b82f6','TMA':'#14b8a6','TCCHO':'#f97316',
    'TPE':'#dc2626','TBD':'#06b6d4','TBS':'#22c55e','TAS':'#1e293b','TIR':'#f59e0b',
    'TMI':'#8b5cf6','TTPIC':'#ef4444','TTPIP':'#9333ea','TEAC':'#eab308',
}
MES_DIA = {'MARZO':'03-10','ABRIL':'04-07','MAYO':'05-05','JUNIO':'06-09','JULIO':'07-07',
    'AGOSTO':'08-04','SEPTIEMBRE':'09-08','OCTUBRE':'10-06','NOVIEMBRE':'11-03','DICIEMBRE':'12-08'}
ING_MAP = {
    'CHUCRUT ENCURTIDO 680 GR (FRASCO)':'CHUCRUT ENCURTIDO','LECHE FRESCA REBECA':'LECHE FRESCA',
    'MALVAVISCOS 240 GR':'MALVAVISCOS','SAL DE CURA #1':'SAL DE CURA',
    'VASOS ACRILICOS DE DEGUSTACION 2 OZ C/TAPA':'VASOS ACRILICOS DE DEGUSTACION 2 OZ',
    'VASOS DE DEGUSTACION 3 OZ C/TAPA':'VASOS DE DEGUSTACION 3 OZ','PAPA':'PAPAS',
}
ING_CATS = {
    'CUAJO':'LACTEO','CUCHARA DE MADERA':'ARTICULO COCINA','MANJAR NESTLE':'LACTEO',
    'SACAROSA':'ENDULZANTE','POLLO PECHUGA ENTERA UN':'POLLO','JUGO DE NARANJA SELECCION':'JUGO',
    'BROCHETA DE BAMBU 10 CM 100 UNI':'ARTICULO COCINA','MASA FILO 320GR':'MASA','APIO RAMA':'FRUTA VERDURA',
}
PROV_NORM = {'lider':'Lider','jumbo':'Jumbo','bidfood':'Bidfood','caserita':'Caserita','gourmet':'Gourmet'}

def clean_cat(s):
    if s is None: return None
    if isinstance(s,float) and (pd.isna(s) or math.isnan(s)): return None
    t = str(s).strip().upper()
    if t in ('NAN','NONE','NULL','#N/D','#N/A',''): return None
    return t.replace('LACTEOS','LACTEO').replace('CONGELADOS','CONGELADO')

# Overrides de paquetes: ingredientes cuyo Neto es por paquete de N unidades
PACK_OVERRIDE = {'BATON 500 PETIT PAIN 3,2GR': 500}

def medida_norm(m):
    if m is None or (isinstance(m,float) and pd.isna(m)): return 'UN'
    return str(m).strip().upper()

def convertir_precio(neto, formato, medida, nombre):
    """Convierte (Neto del envase, Formato, Medida) → (costo por unidad base, unidad base kg/L/UN).
    Regla clave: si Neto/Formato > 1000 el Formato ya está en kg/L (no gramos/cc)."""
    if neto is None or pd.isna(neto): return None, 'UN'
    nom = str(nombre).strip().upper()
    if nom in PACK_OVERRIDE:
        return round(neto / PACK_OVERRIDE[nom], 4), 'UN'
    if formato is None or pd.isna(formato) or float(formato) == 0: formato = 1
    base = float(neto) / float(formato)
    m = medida_norm(medida)
    if m in ('KG','GR','G','GRS'):
        if m != 'KG' and base < 1000:   # gramos reales → a kg
            base *= 1000
        return round(base, 4), 'kg'
    if m in ('L','LT','CC','ML'):
        if m in ('CC','ML') and base < 1000:  # cc reales → a L
            base *= 1000
        return round(base, 4), 'L'
    return round(base, 4), 'UN'  # UN, BANDEJA, etc.

def norm_prov(p):
    if p is None or (isinstance(p,float) and pd.isna(p)): return None
    k = str(p).strip().lower()
    return PROV_NORM.get(k, str(p).strip())

# Hoja 1
df1 = pd.read_excel(FILE,'DETALLE RAMO MADRE ',header=None).iloc[2:,1:7].reset_index(drop=True)
df1.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Seccion','N_Alumnos']
df1 = df1.dropna(subset=['Codigo_ID'])
df1['Semestre']=df1['Semestre'].astype(int); df1['Seccion']=df1['Seccion'].astype(int); df1['N_Alumnos']=df1['N_Alumnos'].astype(int)
# Hoja 2
df2 = pd.read_excel(FILE,'DETALLE DE SUB-RAMOS MENSUALES',header=None).iloc[1:,1:7].reset_index(drop=True)
df2.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Mes','OP']
for c in ['Semestre','Codigo_ID','Siglas','Nombre','Mes']: df2[c]=df2[c].ffill()
df2 = df2.dropna(subset=['OP']); df2['OP']=df2['OP'].astype(str).str.strip(); df2['Mes']=df2['Mes'].astype(str).str.strip().str.upper()
# Hoja 3
df3 = pd.read_excel(FILE,'INGREDIENTES DE SUB-RAMOS',header=None).iloc[1:,1:7].reset_index(drop=True)
df3.columns = ['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3 = df3.dropna(subset=['Ingrediente'])
df3['Cantidad']=pd.to_numeric(df3['Cantidad'],errors='coerce')
df3['Ingrediente']=df3['Ingrediente'].astype(str).str.strip().str.upper()
df3['Sub_Ramo']=df3['Sub_Ramo'].astype(str).str.strip()
# Hoja 4
df4 = pd.read_excel(FILE,'DETALLE PRECIO INGREDIENTES',header=None).iloc[1:,0:8].reset_index(drop=True)
df4.columns = ['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4 = df4.dropna(subset=['Ingrediente'])
df4['Neto']=pd.to_numeric(df4['Neto'],errors='coerce'); df4['Formato']=pd.to_numeric(df4['Formato'],errors='coerce')
df4['Ingrediente']=df4['Ingrediente'].astype(str).str.strip().str.upper()
# Conversión de unidades: costo por unidad base (kg/L/UN) + unidad canónica
_conv = df4.apply(lambda r: convertir_precio(r['Neto'],r['Formato'],r['Medida'],r['Ingrediente']), axis=1)
df4['costo_neto']  = _conv.apply(lambda x: x[0])
df4['unidad_base'] = _conv.apply(lambda x: x[1])
df4['Categoria']=df4['Categoria'].apply(clean_cat)
df4['Proveedor']=df4['Proveedor'].apply(norm_prov)
df4 = df4[df4['Categoria'].notna()]
df4 = df4.drop_duplicates('Ingrediente', keep='first')

# Estructuras
cats_h3 = {clean_cat(c) for c in df3['Categoria'].dropna()} - {None}
cats_h4 = set(df4['Categoria'].dropna()) - {None}
all_cats = sorted(cats_h3 | cats_h4)
prov_map_disp = {}
for p in df4['Proveedor'].dropna().unique():
    k = p.strip().lower()
    if k not in prov_map_disp: prov_map_disp[k] = p.strip()
all_provs = sorted(prov_map_disp.values(), key=str.lower)
talleres = df1.drop_duplicates('Siglas')[['Semestre','Siglas','Nombre']].sort_values('Semestre')
sec_por_tall = df1.groupby('Siglas').apply(lambda g: list(zip(g['Seccion'].astype(int),g['N_Alumnos'].astype(int)))).to_dict()

# Sesiones (cod_clase único: op + mes abreviado + sección, con anti-colisión)
ses_defs = []  # (cod, op_clean, sigla, sec, n_alum, fecha)
cods_usados = set()
for _,row in df2.iterrows():
    sigla=str(row['Siglas']).strip(); op=str(row['OP']).strip(); mes=str(row['Mes']).strip().upper()
    fecha=f"2025-{MES_DIA.get(mes,'06-15')}"; op_clean=re.sub(r'\s+','',op); mes3=mes[:3]
    for (sec,n) in sec_por_tall.get(sigla,[(1,15)]):
        base=f"{op_clean}-{mes3}-S{sec}"; cod=base; k=2
        while cod in cods_usados:
            cod=f"{base}-{k}"; k+=1
        cods_usados.add(cod)
        ses_defs.append((cod, op_clean, sigla, sec, n, fecha))

# Ingredientes por OP (canónicos)
ing_h4_names = set(df4['Ingrediente'].str.upper().str.strip())
op_ings = {}  # op_clean → [(ing_canon, cant_unit, unidad)]
nuevos_ing = {}
for _,row in df3.iterrows():
    op_clean=re.sub(r'\s+','',str(row['Sub_Ramo']).strip())
    ing_raw=row['Ingrediente'].upper().strip(); ing_canon=ING_MAP.get(ing_raw,ing_raw)
    cant=float(row['Cantidad']) if pd.notna(row['Cantidad']) else 0.0
    um=str(row['UM']).strip() if pd.notna(row['UM']) else 'UN'
    if ing_canon not in ing_h4_names and ing_canon not in nuevos_ing:
        nuevos_ing[ing_canon]=(ING_CATS.get(ing_canon) or clean_cat(row['Categoria']) or 'ABARROTE', um)
    op_ings.setdefault(op_clean,{})
    if ing_canon in op_ings[op_clean]:
        c0,u0 = op_ings[op_clean][ing_canon]
        op_ings[op_clean][ing_canon] = (round(c0+cant,5), u0)   # consolidar: sumar cantidades
    else:
        op_ings[op_clean][ing_canon] = (cant, um)
# convertir dict interno a lista de tuplas (ing_canon, cant, um)
op_ings = {op:[(k,v[0],v[1]) for k,v in d.items()] for op,d in op_ings.items()}

print('═'*60)
print('PARSEO EXCEL OK')
print(f'  Categorías:{len(all_cats)}  Proveedores:{len(all_provs)}  Talleres:{len(talleres)}')
print(f'  Ingredientes precio:{len(df4)}  +sin precio:{len(nuevos_ing)}')
print(f'  Sesiones:{len(ses_defs)}  OPs con ingredientes:{len(op_ings)}')
print('═'*60)

# ══════════════════════════════════════════════════════════════════════════
# 0. LIMPIEZA (orden FK)
# ══════════════════════════════════════════════════════════════════════════
print('\n[0] LIMPIEZA DE DATOS DEMO')
delete_all('sesion_ingrediente')
delete_all('sesion')
delete_all('inventario')
delete_all('modulo_taller','modulo_id')   # PK compuesta, libera FK a taller_tipo
delete_all('ingrediente')
delete_all('taller_tipo')
delete_all('proveedor')
delete_all('categoria')

# ══════════════════════════════════════════════════════════════════════════
# 1. CATEGORÍAS
# ══════════════════════════════════════════════════════════════════════════
print('\n[1] CATEGORÍAS')
rows = [{'nombre':c,'color':CAT_COLORS.get(c,'#6b7280')} for c in all_cats]
ins = post('categoria', rows, ret=True)
cat_id = {r['nombre'].upper():r['id'] for r in ins}
print(f'  ✓ {len(ins)} categorías')

# ══════════════════════════════════════════════════════════════════════════
# 2. PROVEEDORES
# ══════════════════════════════════════════════════════════════════════════
print('\n[2] PROVEEDORES')
rows = [{'nombre':p,'activo':True} for p in all_provs]
ins = post('proveedor', rows, ret=True)
prov_id = {r['nombre'].lower():r['id'] for r in ins}
print(f'  ✓ {len(ins)} proveedores')

# ══════════════════════════════════════════════════════════════════════════
# 3. TALLERES
# ══════════════════════════════════════════════════════════════════════════
print('\n[3] TALLERES')
rows = []
for _,row in talleres.iterrows():
    sig=str(row['Siglas']).strip()
    rows.append({'codigo':sig,'nombre':str(row['Nombre']).strip(),
                 'descripcion':f"Semestre {row['Semestre']} | {sig}",'color':TALL_COLORS.get(sig,'#6b7280')})
ins = post('taller_tipo', rows, ret=True)
tall_id = {r['codigo']:r['id'] for r in ins}
print(f'  ✓ {len(ins)} talleres')

# ══════════════════════════════════════════════════════════════════════════
# 4. INGREDIENTES
# ══════════════════════════════════════════════════════════════════════════
print('\n[4] INGREDIENTES')
rows = []
for _,row in df4.iterrows():
    nom=str(row['Ingrediente']).strip().upper()
    rows.append({'nombre':nom,'costo_neto':row['costo_neto'] if pd.notna(row['costo_neto']) else 0,
        'unidad':row['unidad_base'],
        'categoria_id':cat_id.get(row['Categoria']), 'proveedor_id':prov_id.get((row['Proveedor'] or '').lower()),
        'activo':True})
# placeholders sin precio
for nom,(cat,un) in sorted(nuevos_ing.items()):
    rows.append({'nombre':nom,'costo_neto':0,'unidad':un,'categoria_id':cat_id.get(cat),'proveedor_id':None,'activo':True})
ins = post('ingrediente', rows, ret=True, batch=500)
ing_id = {r['nombre'].upper():r['id'] for r in ins}
print(f'  ✓ {len(ins)} ingredientes')

# ══════════════════════════════════════════════════════════════════════════
# 5. SESIONES
# ══════════════════════════════════════════════════════════════════════════
print('\n[5] SESIONES')
rows = []
for (cod,op,sig,sec,n,fecha) in ses_defs:
    tid = tall_id.get(sig)
    if not tid: continue
    rows.append({'cod_clase':cod,'taller_tipo_id':tid,'seccion':sec,'num_alumnos':n,'fecha':fecha})
ins = post('sesion', rows, ret=True, batch=500)
ses_id = {r['cod_clase']:r['id'] for r in ins}
print(f'  ✓ {len(ins)} sesiones')

# ══════════════════════════════════════════════════════════════════════════
# 6. SESION_INGREDIENTE
# ══════════════════════════════════════════════════════════════════════════
print('\n[6] INGREDIENTES POR SESIÓN')
rows = []
faltan_ing = set()
for (cod,op,sig,sec,n,fecha) in ses_defs:
    sid = ses_id.get(cod)
    if not sid: continue
    for (ing_canon,cant,um) in op_ings.get(op,[]):
        iid = ing_id.get(ing_canon.upper())
        if not iid:
            faltan_ing.add(ing_canon); continue
        rows.append({'sesion_id':sid,'ingrediente_id':iid,
            'cantidad_unit':round(cant,5),'cantidad_total':round(cant*n,5),'unidad':um})
print(f'  Total filas a insertar: {len(rows)}')
if faltan_ing:
    print(f'  ⚠ Ingredientes no encontrados ({len(faltan_ing)}): {sorted(faltan_ing)[:10]}')
t0=time.time()
post('sesion_ingrediente', rows, ret=False, batch=2000, on_conflict='sesion_id,ingrediente_id')
print(f'  ✓ {len(rows)} filas insertadas en {time.time()-t0:.1f}s')

# ══════════════════════════════════════════════════════════════════════════
# VERIFICACIÓN
# ══════════════════════════════════════════════════════════════════════════
print('\n' + '═'*60)
print('VERIFICACIÓN FINAL')
def count(t):
    r = requests.get(f'{REST}/{t}', headers={**H,'Prefer':'count=exact','Range':'0-0'})
    return r.headers.get('content-range','?/?').split('/')[-1]
for t in ['categoria','proveedor','taller_tipo','ingrediente','sesion','sesion_ingrediente']:
    print(f'  {t:22} → {count(t)}')
print('═'*60)
