import pandas as pd, sys, re, math
sys.stdout.reconfigure(encoding='utf-8')

FILE = 'Dashboard Culinary Depurado Santiago.xlsx'

# ─── COLORES ──────────────────────────────────────────────────────────────
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
    'TIC':'#ef4444','TIP':'#a855f7','TCCM':'#3b82f6','TMA':'#14b8a6',
    'TCCHO':'#f97316','TPE':'#dc2626','TBD':'#06b6d4','TBS':'#22c55e',
    'TAS':'#1e293b','TIR':'#f59e0b','TMI':'#8b5cf6','TTPIC':'#ef4444',
    'TTPIP':'#9333ea','TEAC':'#eab308',
}
# Mapeo de nombres en Hoja3 → nombre canónico en Hoja4
ING_MAP = {
    'CHUCRUT ENCURTIDO 680 GR (FRASCO)': 'CHUCRUT ENCURTIDO',
    'LECHE FRESCA REBECA':               'LECHE FRESCA',
    'MALVAVISCOS 240 GR':                'MALVAVISCOS',
    'SAL DE CURA #1':                    'SAL DE CURA',
    'VASOS ACRILICOS DE DEGUSTACION 2 OZ C/TAPA': 'VASOS ACRILICOS DE DEGUSTACION 2 OZ',
    'VASOS DE DEGUSTACION 3 OZ C/TAPA': 'VASOS DE DEGUSTACION 3 OZ',
    'PAPA':                              'PAPAS',
}

MES_DIA = {
    'MARZO':'03-10','ABRIL':'04-07','MAYO':'05-05','JUNIO':'06-09',
    'JULIO':'07-07','AGOSTO':'08-04','SEPTIEMBRE':'09-08','OCTUBRE':'10-06',
    'NOVIEMBRE':'11-03','DICIEMBRE':'12-08',
}

def esc(s):
    if s is None or (isinstance(s, float) and (pd.isna(s) or math.isnan(s))):
        return 'NULL'
    return "'" + str(s).replace("'", "''") + "'"

def fnum(v, decimals=4):
    try:
        f = float(v)
        if pd.isna(f) or math.isnan(f): return 'NULL'
        return str(round(f, decimals))
    except: return 'NULL'

def clean_cat(s):
    if s is None: return None
    if isinstance(s, float) and (pd.isna(s) or math.isnan(s)): return None
    t = str(s).strip().upper()
    if t in ('NAN','NONE','NULL','#N/D','#N/A',''): return None
    t = t.replace('LACTEOS','LACTEO').replace('CONGELADOS','CONGELADO')
    return t

# Categorias manuales para ingredientes especiales
ING_CATS = {
    'CUAJO':                          'LACTEO',
    'CUCHARA DE MADERA':              'ARTICULO COCINA',
    'MANJAR NESTLE':                  'LACTEO',
    'SACAROSA':                       'ENDULZANTE',
    'POLLO PECHUGA ENTERA UN':        'POLLO',
    'JUGO DE NARANJA SELECCION':      'JUGO',
    'BROCHETA DE BAMBU 10 CM 100 UNI':'ARTICULO COCINA',
    'MASA FILO 320GR':                'MASA',
    'APIO RAMA':                      'FRUTA VERDURA',
}

def clean_prov(s):
    if not s or str(s).startswith('#'): return None
    s = str(s).strip()
    return s.title() if s.lower() in ('lider','jumbo') else s

# ─── LEER DATOS ──────────────────────────────────────────────────────────
# Hoja 1: Ramos Madre
df1 = pd.read_excel(FILE, sheet_name='DETALLE RAMO MADRE ', header=None)
df1 = df1.iloc[2:, 1:7].reset_index(drop=True)
df1.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Seccion','N_Alumnos']
df1 = df1.dropna(subset=['Codigo_ID'])
df1['Semestre'] = df1['Semestre'].astype(int)
df1['Seccion']  = df1['Seccion'].astype(int)
df1['N_Alumnos']= df1['N_Alumnos'].astype(int)

# Hoja 2: Sub-ramos
df2 = pd.read_excel(FILE, sheet_name='DETALLE DE SUB-RAMOS MENSUALES', header=None)
df2 = df2.iloc[1:, 1:7].reset_index(drop=True)
df2.columns = ['Semestre','Codigo_ID','Siglas','Nombre','Mes','OP']
for c in ['Semestre','Codigo_ID','Siglas','Nombre','Mes']: df2[c] = df2[c].ffill()
df2 = df2.dropna(subset=['OP'])
df2['Semestre'] = df2['Semestre'].astype(int)
df2['OP']       = df2['OP'].astype(str).str.strip()
df2['Mes']      = df2['Mes'].astype(str).str.strip().str.upper()

# Hoja 3: Ingredientes por sub-ramo
df3 = pd.read_excel(FILE, sheet_name='INGREDIENTES DE SUB-RAMOS', header=None)
df3 = df3.iloc[1:, 1:7].reset_index(drop=True)
df3.columns = ['Ramo','Sub_Ramo','Ingrediente','Cantidad','UM','Categoria']
df3 = df3.dropna(subset=['Ingrediente'])
df3['Cantidad']    = pd.to_numeric(df3['Cantidad'], errors='coerce')
df3['Ingrediente'] = df3['Ingrediente'].astype(str).str.strip().str.upper()
df3['Sub_Ramo']    = df3['Sub_Ramo'].astype(str).str.strip()

# Hoja 4: Precios
df4 = pd.read_excel(FILE, sheet_name='DETALLE PRECIO INGREDIENTES', header=None)
df4 = df4.iloc[1:, 0:8].reset_index(drop=True)
df4.columns = ['Categoria','Ingrediente','Proveedor','Cant_Total','Formato','Medida','Neto','Bruto']
df4 = df4.dropna(subset=['Ingrediente'])
df4['Neto']    = pd.to_numeric(df4['Neto'], errors='coerce')
df4['Formato'] = pd.to_numeric(df4['Formato'], errors='coerce')
df4['Ingrediente'] = df4['Ingrediente'].astype(str).str.strip().str.upper()
df4['costo_neto'] = df4.apply(
    lambda r: round(float(r['Neto'])/float(r['Formato']), 4)
    if pd.notna(r['Neto']) and pd.notna(r['Formato']) and float(r['Formato']) > 0
    else None, axis=1
)
# Limpiar categorias y proveedores
df4['Categoria'] = df4['Categoria'].apply(clean_cat)
df4['Proveedor'] = df4['Proveedor'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
df4 = df4[df4['Categoria'].notna()]  # eliminar filas con categoría inválida (#N/D)

# ─── ESTRUCTURAS ─────────────────────────────────────────────────────────
# Categorías únicas (unión Hoja3 + Hoja4, limpiadas)
cats_h3 = {clean_cat(c) for c in df3['Categoria'].dropna()} - {None}
cats_h4 = set(df4['Categoria'].dropna()) - {None}
all_cats = sorted(cats_h3 | cats_h4)

# Normalizar nombres de proveedores conocidos
PROV_NORM = {'lider':'Lider','jumbo':'Jumbo','bidfood':'Bidfood',
             'caserita':'Caserita','gourmet':'Gourmet'}

# Proveedores únicos (desde Hoja 4), dedup por lower, normalizar
provs_raw = df4['Proveedor'].dropna().unique()
prov_map = {}  # lower → display name
for p in provs_raw:
    raw = p.strip()
    k   = raw.lower()
    if k not in prov_map:
        prov_map[k] = PROV_NORM.get(k, raw)
# Normalizar la columna Proveedor en df4 para consistencia en tmp_ing
def norm_prov(p):
    if not p or pd.isna(p): return None
    k = str(p).strip().lower()
    return PROV_NORM.get(k, str(p).strip())
df4['Proveedor'] = df4['Proveedor'].apply(norm_prov)
all_provs = sorted(prov_map.values(), key=str.lower)

# Talleres únicos (Hoja 1)
talleres = df1.drop_duplicates('Siglas')[['Semestre','Siglas','Nombre']].sort_values('Semestre')

# Secciones por taller: {sigla: [(seccion, n_alumnos), ...]}
sec_por_tall = df1.groupby('Siglas').apply(
    lambda g: list(zip(g['Seccion'].astype(int), g['N_Alumnos'].astype(int)))
).to_dict()

# Sesiones (OP × sección)
ses_rows = []  # (cod_clase, op_base, taller_sigla, seccion, n_alumnos, fecha, anio)
for _, row in df2.iterrows():
    sigla = str(row['Siglas']).strip()
    op    = str(row['OP']).strip()
    mes   = str(row['Mes']).strip().upper()
    sem   = int(row['Semestre'])
    anio  = 2025
    fecha_mid = MES_DIA.get(mes, '06-15')
    fecha = f"{anio}-{fecha_mid}"
    # clean cod_clase: replace spaces
    op_clean = re.sub(r'\s+', '', op)  # "SEMANA 1" → "SEMANA1"
    secciones = sec_por_tall.get(sigla, [(1,15)])
    for (sec, n_alum) in secciones:
        cod = f"{op_clean}-S{sec}"
        ses_rows.append((cod, op, sigla, sec, n_alum, fecha))

# Ingredient rows for sesion_ingrediente, applying name mapping
ing_h4_names = set(df4['Ingrediente'].str.upper().str.strip())
si_rows = []        # (op_base, ing_nombre_canonico, cant_unit, unidad)
nuevos_ing = {}     # {nombre: (categoria, unidad)} para ingredientes sin precio

for _, row in df3.iterrows():
    op_clean  = re.sub(r'\s+', '', str(row['Sub_Ramo']).strip())
    ing_raw   = row['Ingrediente'].upper().strip()
    ing_canon = ING_MAP.get(ing_raw, ing_raw)  # aplicar mapeo
    cant      = float(row['Cantidad']) if pd.notna(row['Cantidad']) else 0.0
    um        = str(row['UM']).strip() if pd.notna(row['UM']) else 'UN'
    # Si no existe en Hoja4, agregar como nuevo con precio 0
    if ing_canon not in ing_h4_names and ing_canon not in nuevos_ing:
        cat_raw = ING_CATS.get(ing_canon) or clean_cat(row['Categoria']) or 'ABARROTE'
        nuevos_ing[ing_canon] = (cat_raw, um)
    si_rows.append((op_clean, ing_canon, cant, um))

# ─── GENERAR SQL ─────────────────────────────────────────────────────────
out = []
def L(*args): out.append(' '.join(str(a) for a in args))

L("-- ================================================================")
L("-- CULINARY SANTIAGO – IMPORTACIÓN DATOS REALES")
L("-- Fuente: Dashboard Culinary Depurado Santiago.xlsx")
L("-- Ejecutar en Supabase SQL Editor (New Query)")
L("-- Las tablas deben existir (schema.sql ejecutado previamente)")
L("-- ================================================================")
L("")

# ══ BLOQUE 1: CATEGORÍAS ══════════════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 1: CATEGORÍAS")
L(f"-- {len(all_cats)} categorías del Excel")
L("-- ================================================================")
L("INSERT INTO categoria (nombre, color) VALUES")
rows = [f"  ({esc(c)}, {esc(CAT_COLORS.get(c,'#6b7280'))})" for c in all_cats]
L(",\n".join(rows))
L("ON CONFLICT (nombre) DO UPDATE SET color = EXCLUDED.color;")
L("")

# ══ BLOQUE 2: PROVEEDORES ═════════════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 2: PROVEEDORES")
L(f"-- {len(all_provs)} proveedores del Excel (sin duplicar)")
L("-- ================================================================")
L("INSERT INTO proveedor (nombre, activo)")
L("SELECT v.nombre, true")
L("FROM (VALUES")
rows = [f"  ({esc(p)})" for p in all_provs]
L(",\n".join(rows))
L(") AS v(nombre)")
L("WHERE NOT EXISTS (")
L("  SELECT 1 FROM proveedor")
L("  WHERE UPPER(proveedor.nombre) = UPPER(v.nombre)")
L(");")
L("")

# ══ BLOQUE 3: TALLER_TIPO ════════════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 3: TIPOS DE TALLER (14 ramos reales)")
L("-- ================================================================")
L("INSERT INTO taller_tipo (codigo, nombre, descripcion, color) VALUES")
rows = []
for _, row in talleres.iterrows():
    sig  = str(row['Siglas']).strip()
    nom  = str(row['Nombre']).strip()
    desc = f"Semestre {row['Semestre']} | {sig}"
    col  = TALL_COLORS.get(sig, '#6b7280')
    rows.append(f"  ({esc(sig)}, {esc(nom)}, {esc(desc)}, {esc(col)})")
L(",\n".join(rows))
L("ON CONFLICT (codigo) DO UPDATE SET")
L("  nombre = EXCLUDED.nombre,")
L("  descripcion = EXCLUDED.descripcion,")
L("  color = EXCLUDED.color;")
L("")

# ══ BLOQUE 4: INGREDIENTES ═══════════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 4: INGREDIENTES Y PRECIOS")
L(f"-- {len(df4)} ingredientes (costo_neto = Neto / Formato por unidad de medida)")
L("-- ================================================================")
L("")
L("-- Tabla temporal para upsert masivo")
L("CREATE TEMP TABLE IF NOT EXISTS tmp_ing (")
L("  nombre TEXT, costo_neto DECIMAL(12,4), unidad TEXT,")
L("  cat_nombre TEXT, prov_nombre TEXT")
L(");")
L("TRUNCATE tmp_ing;")
L("")

# Insert en lotes de 100
BATCH = 100
rows_ing = list(df4.iterrows())
for i in range(0, len(rows_ing), BATCH):
    batch = rows_ing[i:i+BATCH]
    L(f"INSERT INTO tmp_ing VALUES")
    vals = []
    for _, row in batch:
        nom   = str(row['Ingrediente']).strip().upper()
        un    = str(row['Medida']).strip() if pd.notna(row['Medida']) else 'UN'
        cost  = row['costo_neto']
        cat   = row['Categoria']
        prov  = str(row['Proveedor']).strip() if pd.notna(row['Proveedor']) else None
        vals.append(f"  ({esc(nom)}, {fnum(cost,4)}, {esc(un)}, {esc(cat)}, {esc(prov)})")
    L(",\n".join(vals) + ";")
    L("")

L("-- Insertar nuevos ingredientes")
L("INSERT INTO ingrediente (nombre, costo_neto, unidad, categoria_id, proveedor_id, activo, fecha_actualizacion)")
L("SELECT t.nombre, t.costo_neto, t.unidad, c.id, p.id, true, CURRENT_DATE")
L("FROM tmp_ing t")
L("LEFT JOIN categoria c ON UPPER(c.nombre) = UPPER(t.cat_nombre)")
L("LEFT JOIN proveedor  p ON UPPER(p.nombre) = UPPER(t.prov_nombre)")
L("WHERE NOT EXISTS (")
L("  SELECT 1 FROM ingrediente WHERE UPPER(ingrediente.nombre) = UPPER(t.nombre)")
L(");")
L("")
L("-- Actualizar ingredientes existentes con precios correctos")
L("UPDATE ingrediente SET")
L("  costo_neto           = t.costo_neto,")
L("  unidad               = t.unidad,")
L("  categoria_id         = c.id,")
L("  proveedor_id         = p.id,")
L("  fecha_actualizacion  = CURRENT_DATE")
L("FROM tmp_ing t")
L("LEFT JOIN categoria c ON UPPER(c.nombre) = UPPER(t.cat_nombre)")
L("LEFT JOIN proveedor  p ON UPPER(p.nombre) = UPPER(t.prov_nombre)")
L("WHERE UPPER(ingrediente.nombre) = UPPER(t.nombre);")
L("")
L("DROP TABLE IF EXISTS tmp_ing;")
L("")

# ══ BLOQUE 4B: INGREDIENTES SIN PRECIO ═══════════════════════════════════
if nuevos_ing:
    L("-- ================================================================")
    L(f"-- BLOQUE 4B: {len(nuevos_ing)} INGREDIENTES EN SESIONES SIN PRECIO EXACTO")
    L("-- costo_neto = 0 → requieren actualizacion manual de precio")
    L("-- ================================================================")
    L("INSERT INTO ingrediente (nombre, costo_neto, unidad, categoria_id, activo, fecha_actualizacion)")
    L("SELECT v.nombre, 0, v.unidad, c.id, true, CURRENT_DATE")
    L("FROM (VALUES")
    rows = []
    for nom, (cat, un) in sorted(nuevos_ing.items()):
        rows.append(f"  ({esc(nom)}, {esc(un)}, {esc(cat)})")
    L(",\n".join(rows))
    L(") AS v(nombre, unidad, cat_nombre)")
    L("LEFT JOIN categoria c ON UPPER(c.nombre) = UPPER(v.cat_nombre)")
    L("WHERE NOT EXISTS (")
    L("  SELECT 1 FROM ingrediente WHERE UPPER(ingrediente.nombre) = UPPER(v.nombre)")
    L(");")
    L("")

# ══ BLOQUE 5: SESIONES ═══════════════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 5: SESIONES DE CLASE")
L(f"-- {len(ses_rows)} sesiones (sub-ramos × secciones por taller)")
L("-- ================================================================")
L("")
L("CREATE TEMP TABLE IF NOT EXISTS tmp_ses (")
L("  cod_clase TEXT, op_base TEXT, taller_codigo TEXT,")
L("  seccion INT, num_alumnos INT, fecha DATE")
L(");")
L("TRUNCATE tmp_ses;")
L("")

BATCH_SES = 200
for i in range(0, len(ses_rows), BATCH_SES):
    batch = ses_rows[i:i+BATCH_SES]
    L(f"INSERT INTO tmp_ses VALUES")
    vals = []
    for (cod, op, sig, sec, n, fecha) in batch:
        op_c = re.sub(r'\s+','',op)
        vals.append(f"  ({esc(cod)}, {esc(op_c)}, {esc(sig)}, {sec}, {n}, {esc(fecha)})")
    L(",\n".join(vals) + ";")
    L("")

L("-- Insertar solo sesiones nuevas (por cod_clase)")
L("INSERT INTO sesion (cod_clase, taller_tipo_id, seccion, num_alumnos, fecha)")
L("SELECT t.cod_clase, tt.id, t.seccion, t.num_alumnos, t.fecha")
L("FROM tmp_ses t")
L("JOIN taller_tipo tt ON tt.codigo = t.taller_codigo")
L("WHERE NOT EXISTS (")
L("  SELECT 1 FROM sesion WHERE sesion.cod_clase = t.cod_clase")
L(");")
L("")

# ══ BLOQUE 6: SESION_INGREDIENTE ═════════════════════════════════════════
L("-- ================================================================")
L("-- BLOQUE 6: INGREDIENTES POR SESIÓN")
L(f"-- {len(si_rows)} filas base → multiplicadas automáticamente por secciones")
L("-- cantidad_total = cantidad_unit × num_alumnos (calculado en SQL)")
L("-- ================================================================")
L("")
L("CREATE TEMP TABLE IF NOT EXISTS tmp_si (")
L("  op_base TEXT, ing_nombre TEXT, cant_unit DECIMAL(12,5), unidad TEXT")
L(");")
L("TRUNCATE tmp_si;")
L("")

BATCH_SI = 300
for i in range(0, len(si_rows), BATCH_SI):
    batch = si_rows[i:i+BATCH_SI]
    L(f"INSERT INTO tmp_si VALUES")
    vals = []
    for (op, ing, cant, un) in batch:
        vals.append(f"  ({esc(op)}, {esc(ing)}, {fnum(cant,5)}, {esc(un)})")
    L(",\n".join(vals) + ";")
    L("")

L("-- Insertar/actualizar ingredientes de sesión")
L("-- tmp_ses vincula op_base con cada sesion (cod_clase)")
L("INSERT INTO sesion_ingrediente (sesion_id, ingrediente_id, cantidad_unit, cantidad_total, unidad)")
L("SELECT")
L("  s.id                                  AS sesion_id,")
L("  i.id                                  AS ingrediente_id,")
L("  si.cant_unit                          AS cantidad_unit,")
L("  ROUND(si.cant_unit * s.num_alumnos, 5) AS cantidad_total,")
L("  si.unidad")
L("FROM tmp_si si")
L("JOIN tmp_ses ts ON ts.op_base = si.op_base")
L("JOIN sesion  s  ON s.cod_clase = ts.cod_clase")
L("JOIN ingrediente i ON UPPER(i.nombre) = UPPER(si.ing_nombre)")
L("ON CONFLICT (sesion_id, ingrediente_id) DO UPDATE SET")
L("  cantidad_unit  = EXCLUDED.cantidad_unit,")
L("  cantidad_total = EXCLUDED.cantidad_total,")
L("  unidad         = EXCLUDED.unidad;")
L("")
L("DROP TABLE IF EXISTS tmp_ses;")
L("DROP TABLE IF EXISTS tmp_si;")
L("")
L("-- ================================================================")
L("-- VERIFICACIÓN FINAL")
L("-- ================================================================")
L("SELECT")
L("  (SELECT COUNT(*) FROM categoria)         AS categorias,")
L("  (SELECT COUNT(*) FROM proveedor)          AS proveedores,")
L("  (SELECT COUNT(*) FROM taller_tipo)        AS talleres,")
L("  (SELECT COUNT(*) FROM ingrediente)        AS ingredientes,")
L("  (SELECT COUNT(*) FROM sesion)             AS sesiones,")
L("  (SELECT COUNT(*) FROM sesion_ingrediente) AS sesion_ingredientes;")
L("")
L("-- ================================================================")
L("-- DIAGNÓSTICO: Ingredientes sin sesión (solo en maestro, no en clases)")
L("-- ================================================================")
L("SELECT i.nombre, i.unidad, i.costo_neto, c.nombre AS categoria")
L("FROM ingrediente i")
L("LEFT JOIN sesion_ingrediente si ON si.ingrediente_id = i.id")
L("LEFT JOIN categoria c ON c.id = i.categoria_id")
L("WHERE si.id IS NULL")
L("ORDER BY i.nombre;")
L("")
L("-- ================================================================")
L("-- RESUMEN DE COSTOS POR TALLER (vista rápida)")
L("-- ================================================================")
L("SELECT")
L("  tt.codigo, tt.nombre,")
L("  COUNT(DISTINCT s.id)   AS total_sesiones,")
L("  SUM(s.num_alumnos)     AS total_alumnos,")
L("  ROUND(SUM(si.cantidad_total * i.costo_neto)::DECIMAL / NULLIF(SUM(s.num_alumnos),0), 0)")
L("    AS costo_prom_alumno_clp")
L("FROM taller_tipo tt")
L("LEFT JOIN sesion s ON s.taller_tipo_id = tt.id")
L("LEFT JOIN sesion_ingrediente si ON si.sesion_id = s.id")
L("LEFT JOIN ingrediente i ON i.id = si.ingrediente_id")
L("GROUP BY tt.id, tt.codigo, tt.nombre")
L("ORDER BY costo_prom_alumno_clp DESC NULLS LAST;")
L("")

# ─── GUARDAR ─────────────────────────────────────────────────────────────
sql_text = "\n".join(out)
with open('import_culinary.sql', 'w', encoding='utf-8') as f:
    f.write(sql_text)

print("SQL generado: import_culinary.sql")
print(f"  Categorias:           {len(all_cats)}")
print(f"  Proveedores:          {len(all_provs)}")
print(f"  Talleres:             {len(talleres)}")
print(f"  Ingredientes:         {len(df4)}")
print(f"  Sesiones (OP×sec):   {len(ses_rows)}")
print(f"  SI rows base:         {len(si_rows)}")
print(f"  Lineas SQL:           {len(out)}")
print(f"  Tamaño:               {len(sql_text)//1024} KB")
