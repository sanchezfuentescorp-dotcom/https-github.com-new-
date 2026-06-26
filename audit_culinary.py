import sys, re
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')
with open('index.html', encoding='utf-8') as f:
    content = f.read()
lines = content.splitlines()

issues = []
ok = []

# ── 1. FUNCIONES DUPLICADAS ────────────────────────────────────────────
func_lines = {}
func_counts = Counter()
for i, l in enumerate(lines, 1):
    s = l.strip()
    if s.startswith('function ') and '(' in s:
        name = s.split('(')[0].replace('function ', '').strip()
        func_counts[name] += 1
        if name not in func_lines:
            func_lines[name] = []
        func_lines[name].append(i)

dups = [(n, c, func_lines[n]) for n, c in func_counts.items() if c > 1]
if dups:
    for n, c, lns in sorted(dups):
        issues.append(f'[DUP] {n} definida {c}x en lineas {lns}')
else:
    ok.append('Sin funciones duplicadas')

# ── 2. REFERENCIAS CIRCULARES _orig ───────────────────────────────────
circ = re.findall(r'const\s+_\w+\s*=\s*(\w+)', content)
for fn in circ:
    if fn in func_counts:
        issues.append(f'[CIRC] const _x = {fn} — posible referencia circular activa')

if not circ:
    ok.append('Sin referencias circulares _orig')

# ── 3. FUNCIONES onclick NO DEFINIDAS ─────────────────────────────────
defined_fns = set(func_counts.keys())
known_globals = {
    'nav', 'toggleGroup', 'toggleDark', 'openModal', 'closeModal', 'toast',
    'GET', 'POST', 'PATCH', 'DEL', 'DELW', 'clp', 'v', 'hexAlpha',
    'loadAll', 'loadDash', 'loadSesiones', 'loadIngredientes', 'loadProveedores',
    'loadTalleres', 'loadCarreras', 'loadModulos', 'loadSemanas', 'loadInventario',
    'loadCompras', 'loadOrdenCompra', 'loadHistorial', 'loadPresupuesto',
    'loadSimulador', 'loadComparativa', 'loadBackup', 'loadUsuarios',
    'loadCentroCosto', 'loadCategorias', 'loadConfig', 'initDark',
    'renderKPIs', 'renderIngredientes', 'renderSesFiltered', 'renderCalendario',
    'openSesModal', 'saveSesion', 'openIngModal', 'saveIngrediente',
    'openProvModal', 'saveProveedor', 'openTallModal', 'saveTaller',
    'openCarModal', 'saveCarrera', 'openModuloModal', 'saveModulo',
    'openSemanaModal', 'saveSemana', 'openInvModal', 'saveInventario',
    'openCCModal', 'saveCentroCC', 'openCatModal', 'saveCategoria',
    'openUsuarioModal', 'saveUsuario',
    'addSesIngRow', 'addModTallRow', 'filterTable',
    'verDetalleSesion', 'duplicarSesion', 'eliminarSesion', 'editarSesion',
    'verDetalleProv', 'verDetalleIng', 'editarIng', 'eliminarIng',
    'ocNuevaOC', 'ocShowLista', 'ocVerDetalle', 'ocEmitirOC', 'ocAnularOC',
    'ocPDFOC', 'ocAbrirRecepcion', 'ocGuardarRecepcion', 'ocEnviarCorreo',
    'ocAddItemManual', 'ocFilterItems', 'ocAddFromItems', 'ocAutoFillProv',
    'ocSaveDraft', 'ocRemoveItem', 'renderOCLista', 'renderOCKpis',
    'crearOCdesdeListaCompras', 'exportOCExcel', 'exportComprasExcel',
    'abrirInformeCarrera', 'cerrarInformeCarrera',
    'abrirBusqueda', 'cerrarBusqueda', 'buscarGlobal', 'gsKeyNav',
    'exportDashPDF', 'exportDashExcel', 'exportCierreMensual',
    'generarPlantilla', 'handleFile', 'ejecutarImport', 'showImpTab', 'resetImport',
    'handlePreciosFile', 'ejecutarImportPrecios', 'resetPreciosImport',
    'populateDashFilters', 'limpiarFiltrosSes', 'limpiarFiltrosProv',
    'renderAlertTable', 'renderUltimasSesiones', 'renderTopSesiones',
    'sortTableBy', 'guardarComoPlantilla', 'openPlantillasModal',
    'cargarPlantilla', 'borrarPlantilla', 'recalcItemsRec',
    'refreshNotifs', 'toggleNotifPanel', 'cerrarNotif',
    'toggleSesVista', 'calPrevMes', 'calNextMes',
    'ocVerVersiones', 'renderOCVersiones', 'verHistorialProv', 'verHistorialPrecios',
    'renderSimulador', 'calcSimulador', 'exportSimulador',
    'editPresupuesto', 'savePresupuesto', 'backupData', 'restoreBackup',
    'exportBackup', 'cfgSave', 'cfgClose', 'alert', 'confirm', 'prompt',
    'parseInt', 'parseFloat', 'String', 'Number', 'Boolean', 'Array',
    'Object', 'JSON', 'Math', 'Date', 'console', 'window', 'document',
    'setTimeout', 'setInterval', 'clearInterval', 'fetch', 'localStorage',
    'encodeURIComponent', 'decodeURIComponent',
}
all_known = defined_fns | known_globals

# Extract all event handler calls
handler_pattern = re.compile(r'on(?:click|input|change|focus|submit|keydown|dragover|dragleave|drop)=["\']([^"\']+)["\']')
missing_fns = set()
for match in handler_pattern.finditer(content):
    call_str = match.group(1)
    for fn_match in re.finditer(r'(\b[a-z][a-zA-Z0-9_]+)\s*\(', call_str):
        fn = fn_match.group(1)
        if fn not in all_known and len(fn) > 3:
            missing_fns.add(fn)

if missing_fns:
    for fn in sorted(missing_fns):
        issues.append(f'[UNDEF] {fn}() llamada desde event handler pero no definida')
else:
    ok.append('Todas las funciones en event handlers estan definidas')

# ── 4. IDs REFERENCIADOS EN JS QUE NO EXISTEN EN HTML ─────────────────
html_ids = set(re.findall(r'\bid=["\']([^"\']+)["\']', content))
js_getelem = re.findall(r"getElementById\(['\"]([^'\"]+)['\"]\)", content)

# IDs generados dinamicamente (contienen variable/patron)
dynamic_prefixes = ['rec-ico-', 'rec-item-', 'ver-', 'ing-ph-', 'ph-']
missing_ids = set()
for eid in js_getelem:
    if eid not in html_ids:
        if not any(eid.startswith(p) for p in dynamic_prefixes):
            missing_ids.add(eid)

if missing_ids:
    for eid in sorted(missing_ids):
        issues.append(f'[ID] getElementById("{eid}") — no existe en HTML')
else:
    ok.append('Todos los getElementById apuntan a IDs existentes en HTML')

# ── 5. FUNCIONES NUEVAS DEFINIDAS Y SUS IDs CLAVE ────────────────────
new_fns = [
    'sortTableBy', 'renderTopSesiones', 'guardarComoPlantilla',
    'openPlantillasModal', 'cargarPlantilla', 'borrarPlantilla',
    'ocAbrirRecepcion', 'recalcItemsRec', 'ocGuardarRecepcion',
    'exportCierreMensual', 'handlePreciosFile', 'ejecutarImportPrecios',
    'resetPreciosImport'
]
for fn in new_fns:
    if fn in defined_fns:
        ok.append(f'{fn}() definida correctamente')
    else:
        issues.append(f'[MISSING] {fn}() no esta definida')

# IDs clave de las nuevas features
key_ids = [
    'dash-top-ses', 'plantillas-lista', 'ov-ses-plantillas',
    'rec-items-wrap', 'rec-items-lista', 'rec-items-resumen',
    'imp-precios', 'precios-drop', 'precios-preview', 'precios-kpis',
    'precios-preview-tbody', 'precios-file-input', 'precios-status',
]
for eid in key_ids:
    if eid in html_ids:
        ok.append(f'ID #{eid} existe en HTML')
    else:
        issues.append(f'[MISSING_ID] #{eid} no encontrado en HTML')

# ── 6. RESUMEN ─────────────────────────────────────────────────────────
print('=' * 60)
print('AUDITORIA CULINARY COSTOS')
print('=' * 60)
print(f'Lineas totales: {len(lines)}')
print(f'Funciones JS definidas: {len(func_counts)}')
print(f'IDs HTML encontrados: {len(html_ids)}')
print()

print(f'RESULTADOS OK ({len(ok)}):')
for o in ok:
    print(f'  [OK] {o}')

print()
if issues:
    print(f'PROBLEMAS ENCONTRADOS ({len(issues)}):')
    for iss in issues:
        print(f'  {iss}')
    sys.exit(1)
else:
    print('RESULTADO FINAL: LIMPIO — 0 problemas detectados')
