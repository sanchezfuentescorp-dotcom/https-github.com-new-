import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, pandas as pd, sys, re, collections, statistics
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}
FILE=r'C:\Users\dmoov\Downloads\Dashboard Culinary Depurado Santiago.xlsx'
def getall(t,sel='*'):
    out,off=[],0
    while True:
        r=requests.get(f'{URL}/rest/v1/{t}?select={sel}',headers={**H,'Range':f'{off}-{off+999}'})
        d=r.json()
        if not isinstance(d,list) or not d: break
        out+=d
        if len(d)<1000: break
        off+=1000
    return out
def clp(v): return ('$'+f'{v:,.0f}').replace(',','.')
OK='✓'; BAD='✗ ERROR'; W='⚠'
problemas=[]

print('═'*72); print('AUDITORÍA PROFUNDA — CULINARY SANTIAGO'); print('═'*72)

# Cargar datos base
ings={i['id']:i for i in getall('ingrediente','id,nombre,costo_neto,unidad,categoria_id,proveedor_id')}
ses={s['id']:s for s in getall('sesion','id,cod_clase,taller_tipo_id,num_alumnos,fecha,semana_id')}
si=getall('sesion_ingrediente','id,sesion_id,ingrediente_id,cantidad_unit,cantidad_total,unidad')
cats={c['id']:c['nombre'] for c in getall('categoria','id,nombre')}

# ── 1. MODELO DE COSTEO: cantidad_total == cantidad_unit en TODAS las filas ──
print('\n[1] MODELO DE COSTEO (cantidad_total = cantidad de receta)')
dif=[r for r in si if abs((r['cantidad_total'] or 0)-(r['cantidad_unit'] or 0))>1e-9]
print(f'  {OK if not dif else BAD} filas con cantidad_total ≠ cantidad_unit: {len(dif)}')
if dif: problemas.append(f'{len(dif)} filas con cantidad_total≠cantidad_unit')

# ── 2. v_costo_sesion: verificación manual ──
print('\n[2] v_costo_sesion (costo_total y costo_por_alumno)')
vsesion={s['id']:s for s in getall('v_costo_sesion','id,costo_total,num_alumnos,costo_por_alumno')}
man=collections.defaultdict(float)
for r in si: man[r['sesion_id']]+=(r['cantidad_total'] or 0)*(ings.get(r['ingrediente_id'],{}).get('costo_neto') or 0)
err_ct=err_cpa=0
for sid,vw in vsesion.items():
    if abs((vw['costo_total'] or 0)-round(man[sid],2))>0.5: err_ct+=1
    n=vw['num_alumnos'] or 0
    esp_cpa=round(man[sid]/n,2) if n>0 else 0
    if abs((vw['costo_por_alumno'] or 0)-esp_cpa)>0.5: err_cpa+=1
print(f'  {OK if not err_ct else BAD} costo_total: {err_ct} discrepancias / {len(vsesion)}')
print(f'  {OK if not err_cpa else BAD} costo_por_alumno: {err_cpa} discrepancias')
if err_ct: problemas.append(f'v_costo_sesion costo_total {err_ct} errores')
if err_cpa: problemas.append(f'v_costo_sesion costo_por_alumno {err_cpa} errores')

# ── 3. v_costo_categoria ──
print('\n[3] v_costo_categoria')
vcat={c['categoria']:c for c in getall('v_costo_categoria','categoria,costo_total,num_ingredientes')}
man_cat=collections.defaultdict(float)
for r in si:
    ing=ings.get(r['ingrediente_id']);
    if not ing: continue
    cat=cats.get(ing['categoria_id'])
    man_cat[cat]+=(r['cantidad_total'] or 0)*(ing['costo_neto'] or 0)
errc=0
for cat,vw in vcat.items():
    if abs((vw['costo_total'] or 0)-round(man_cat.get(cat,0),2))>0.5: errc+=1
print(f'  {OK if not errc else BAD} costo_total por categoría: {errc} discrepancias / {len(vcat)}')
tot_cat=sum(v['costo_total'] or 0 for v in vcat.values())
print(f'  Suma categorías = {clp(tot_cat)}')
if errc: problemas.append(f'v_costo_categoria {errc} errores')

# ── 4. v_top_ingredientes ──
print('\n[4] v_top_ingredientes')
vtop={t['id']:t for t in getall('v_top_ingredientes','id,costo_total,cantidad_usada')}
man_ing=collections.defaultdict(lambda:[0.0,0.0])
for r in si:
    man_ing[r['ingrediente_id']][0]+=(r['cantidad_total'] or 0)
    man_ing[r['ingrediente_id']][1]+=(r['cantidad_total'] or 0)*(ings.get(r['ingrediente_id'],{}).get('costo_neto') or 0)
errt=0
for iid,vw in vtop.items():
    if abs((vw['costo_total'] or 0)-round(man_ing[iid][1],2))>0.5: errt+=1
print(f'  {OK if not errt else BAD} costo_total por ingrediente: {errt} discrepancias / {len(vtop)}')
if errt: problemas.append(f'v_top_ingredientes {errt} errores')

# ── 5. v_costo_taller (bug conocido) ──
print('\n[5] v_costo_taller')
vt=getall('v_costo_taller','codigo,total_alumnos,costo_total,costo_promedio_alumno')
real_alu=collections.defaultdict(int)
tmap={t['id']:t['codigo'] for t in getall('taller_tipo','id,codigo')}
for s in ses.values(): real_alu[tmap[s['taller_tipo_id']]]+=s['num_alumnos']
bug=sum(1 for t in vt if (t['total_alumnos'] or 0)!=real_alu.get(t['codigo'],0))
print(f'  {BAD if bug else OK} talleres con total_alumnos inflado: {bug}/{len(vt)}  (fix_v_costo_taller.sql pendiente)')
if bug: problemas.append(f'v_costo_taller fan-out en {bug} talleres (fix SQL pendiente)')

# ── 6. v_costo_modulo / carrera (dependientes + modulo_taller) ──
print('\n[6] v_costo_modulo / v_costo_carrera')
mt=getall('modulo_taller','modulo_id')
print(f'  {W if not mt else OK} modulo_taller: {len(mt)} vínculos' + (' (vacío → módulos/carreras en $0)' if not mt else ''))
if not mt: problemas.append('modulo_taller vacío → v_costo_modulo/carrera en $0 (no había datos en Excel)')

# ── 7. OUTLIERS DE CANTIDAD ──
print('\n[7] OUTLIERS DE CANTIDAD (posibles errores de dato del Excel)')
bying=collections.defaultdict(list)
for r in si:
    if (r['cantidad_unit'] or 0)>0: bying[r['ingrediente_id']].append(r['cantidad_unit'])
outliers=[]
for iid,cs in bying.items():
    if len(cs)<3: continue
    med=statistics.median(cs)
    if med>0 and max(cs)>med*20:
        imp=sum((r['cantidad_total'] or 0)*(ings[iid]['costo_neto'] or 0) for r in si if r['ingrediente_id']==iid and (r['cantidad_unit'] or 0)>med*20)
        outliers.append((ings[iid]['nombre'],med,max(cs),imp))
outliers.sort(key=lambda x:-x[3])
print(f'  {W if outliers else OK} ingredientes con cantidad atípica (>20× mediana): {len(outliers)}')
for n,med,mx,imp in outliers[:6]: print(f'      {n[:34]:36} mediana={med:.4f} max={mx:.4f} impacto={clp(imp)}')

# ── 8. PRECIOS vs EXCEL ──
print('\n[8] PRECIOS: ingredientes con $0')
p0=[i for i in ings.values() if not i['costo_neto']]
print(f'  {W} ingredientes con precio $0: {len(p0)}  ({", ".join(i["nombre"][:18] for i in p0[:6])}...)')

# ── 9. CONTEOS Y AÑO ──
print('\n[9] CONTEOS / AÑO / INTEGRIDAD')
print(f'  Ingredientes:{len(ings)}  Sesiones:{len(ses)}  Det:{len(si)}')
anios=collections.Counter((s['fecha'] or '')[:4] for s in ses.values())
print(f'  {OK if set(anios)=={"2026"} else BAD} años sesiones: {dict(anios)}')
huerf=sum(1 for r in si if r['ingrediente_id'] not in ings or r['sesion_id'] not in ses)
print(f'  {OK if not huerf else BAD} sesion_ingrediente huérfanos: {huerf}')
sem=getall('semana','anio');
print(f'  {OK} semanas: {len(sem)} (años {sorted(set(s["anio"] for s in sem))})')

# ── 10. GASTO TOTAL ──
print('\n[10] GASTO MATERIA PRIMA')
gt=sum((r['cantidad_total'] or 0)*(ings.get(r['ingrediente_id'],{}).get('costo_neto') or 0) for r in si)
print(f'  Total: {clp(gt)}')
print(f'  = suma v_costo_sesion: {clp(sum(v["costo_total"] or 0 for v in vsesion.values()))}')
print(f'  = suma v_costo_categoria: {clp(tot_cat)}')

print('\n'+'═'*72)
if problemas:
    print(f'RESUMEN: {len(problemas)} punto(s) de atención:')
    for p in problemas: print(f'  • {p}')
else:
    print('RESUMEN: sin errores detectados ✓')
print('═'*72)
