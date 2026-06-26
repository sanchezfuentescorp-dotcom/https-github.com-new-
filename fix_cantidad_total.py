import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys, json, time
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
REST=f'{URL}/rest/v1'
H={'apikey':KEY,'Authorization':f'Bearer {KEY}','Content-Type':'application/json'}
def getall(t,sel):
    out,off=[],0
    while True:
        r=requests.get(f'{REST}/{t}?select={sel}',headers={**H,'Range':f'{off}-{off+999}'})
        d=r.json()
        if not isinstance(d,list) or not d: break
        out+=d
        if len(d)<1000: break
        off+=1000
    return out
def clp(v): return ('$'+f'{v:,.0f}').replace(',','.')

print('Leyendo sesion_ingrediente...')
rows=getall('sesion_ingrediente','sesion_id,ingrediente_id,cantidad_unit,unidad')
print(f'  {len(rows)} filas')

# Reconstruir con cantidad_total = cantidad_unit (modelo por clase/sección)
nuevas=[{'sesion_id':r['sesion_id'],'ingrediente_id':r['ingrediente_id'],
         'cantidad_unit':r['cantidad_unit'],'cantidad_total':r['cantidad_unit'],
         'unidad':r['unidad']} for r in rows]

print('Borrando sesion_ingrediente...')
rd=requests.delete(f'{REST}/sesion_ingrediente?id=gt.0',headers={**H,'Prefer':'return=minimal'})
print(f'  DELETE [{rd.status_code}]')

print('Re-insertando con cantidad_total = cantidad de receta...')
t0=time.time()
for i in range(0,len(nuevas),2000):
    chunk=nuevas[i:i+2000]
    r=requests.post(f'{REST}/sesion_ingrediente',headers={**H,'Prefer':'return=minimal'},data=json.dumps(chunk))
    if r.status_code not in (200,201,204):
        print(f'  ✗ ERROR [{i}] {r.status_code}: {r.text[:200]}'); sys.exit(1)
print(f'  ✓ {len(nuevas)} filas en {time.time()-t0:.1f}s')

# Verificar nuevo monto
ings={i['id']:(i['costo_neto'] or 0) for i in getall('ingrediente','id,costo_neto')}
si=getall('sesion_ingrediente','ingrediente_id,cantidad_total')
tot=sum((r['cantidad_total'] or 0)*ings.get(r['ingrediente_id'],0) for r in si)
print(f'\nNUEVO GASTO MAT. PRIMA: {clp(tot)}')
print(f'  (antes: $130.397.525)')
