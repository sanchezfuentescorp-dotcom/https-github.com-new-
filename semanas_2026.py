import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys, datetime, collections
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}','Content-Type':'application/json'}
def getall(t,sel):
    out,off=[],0
    while True:
        r=requests.get(f'{URL}/rest/v1/{t}?select={sel}',headers={**H,'Range':f'{off}-{off+999}'})
        d=r.json()
        if not isinstance(d,list) or not d: break
        out+=d
        if len(d)<1000: break
        off+=1000
    return out

# ── 1. CAMBIAR AÑO DE SESIONES 2025 → 2026 ─────────────────────────────
print('[1] CAMBIANDO AÑO DE SESIONES 2025 → 2026')
ses=getall('sesion','fecha')
fechas=sorted({s['fecha'] for s in ses if s['fecha'] and s['fecha'].startswith('2025')})
print(f'    Fechas 2025 distintas: {len(fechas)}')
for f in fechas:
    nueva='2026'+f[4:]
    r=requests.patch(f'{URL}/rest/v1/sesion?fecha=eq.{f}',headers={**H,'Prefer':'return=minimal'},
                     json={'fecha':nueva})
    print(f'    {f} → {nueva}  [{r.status_code}]')

# ── 2. BORRAR SEMANAS DEMO 2025 ────────────────────────────────────────
print('\n[2] ELIMINANDO SEMANAS DEMO 2025')
r=requests.delete(f'{URL}/rest/v1/semana?anio=eq.2025',headers={**H,'Prefer':'return=minimal'})
print(f'    Semanas 2025 eliminadas [{r.status_code}]')

# ── 3. GENERAR CALENDARIO 2026 (lunes-viernes, marzo-diciembre) ────────
print('\n[3] GENERANDO CALENDARIO ACADÉMICO 2026')
MESES={3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',
       9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
# primer lunes de marzo 2026
d=datetime.date(2026,3,2)  # 2026-03-02 es lunes
while d.weekday()!=0: d+=datetime.timedelta(days=1)
filas=[]; num=1
fin_anio=datetime.date(2026,12,31)
while d<=fin_anio:
    viernes=d+datetime.timedelta(days=4)
    sem='1er sem' if d.month<=7 else '2do sem'
    filas.append({'numero':num,'anio':2026,
                  'nombre':f'Semana {num} – {MESES[d.month]} ({sem})',
                  'fecha_inicio':d.isoformat(),'fecha_fin':viernes.isoformat()})
    num+=1; d+=datetime.timedelta(days=7)
r=requests.post(f'{URL}/rest/v1/semana',headers={**H,'Prefer':'return=representation'},json=filas)
if r.status_code in (200,201):
    ins=r.json()
    print(f'    ✓ {len(ins)} semanas creadas (2026-03-02 a {filas[-1]["fecha_fin"]})')
    print(f'    Ejemplos:')
    for s in ins[:3]+ins[-2:]:
        print(f'      Sem {s["numero"]:>2}: {s["fecha_inicio"]} a {s["fecha_fin"]}  {s["nombre"]}')
else:
    print(f'    ✗ ERROR [{r.status_code}]: {r.text[:300]}')

# ── 4. VERIFICACIÓN ────────────────────────────────────────────────────
print('\n[4] VERIFICACIÓN')
ses2=getall('sesion','fecha')
anios=collections.Counter((s['fecha'] or '')[:4] for s in ses2)
print(f'    Sesiones por año: {dict(anios)}')
sem=getall('semana','numero,anio')
print(f'    Semanas en BD: {len(sem)}  (años: {sorted({s["anio"] for s in sem})})')
