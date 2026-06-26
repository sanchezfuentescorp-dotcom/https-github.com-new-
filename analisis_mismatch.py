import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}

def getall(t,sel):
    out,off=[],0
    while True:
        r=requests.get(f'{URL}/rest/v1/{t}?select={sel}',headers={**H,'Range':f'{off}-{off+999}'})
        d=r.json();
        if not isinstance(d,list) or not d: break
        out+=d
        if len(d)<1000: break
        off+=1000
    return out

ings={i['id']:i for i in getall('ingrediente','id,nombre,unidad,costo_neto')}
si=getall('sesion_ingrediente','ingrediente_id,unidad,cantidad_total')

# Agregar impacto por ingrediente con mismatch UM receta vs unidad precio
agg={}
for r in si:
    iid=r['ingrediente_id']; ing=ings.get(iid)
    if not ing: continue
    um_rec=(r['unidad'] or '').strip().lower()
    um_pre=(ing['unidad'] or '').strip().lower()
    if um_rec!=um_pre:
        key=ing['nombre']
        impacto=(r['cantidad_total'] or 0)*(ing['costo_neto'] or 0)
        a=agg.setdefault(key,{'um_rec':um_rec,'um_pre':um_pre,'costo':ing['costo_neto'],'impacto':0,'cant':0})
        a['impacto']+=impacto; a['cant']+=r['cantidad_total'] or 0

print('INGREDIENTES CON MISMATCH UNIDAD (receta ≠ precio) – por impacto $')
print('─'*88)
print(f"{'INGREDIENTE':34} {'um_rec':>7} {'um_pre':>7} {'$/u_precio':>11} {'IMPACTO_$':>14}")
print('─'*88)
tot=0
for k,a in sorted(agg.items(),key=lambda x:-x[1]['impacto']):
    tot+=a['impacto']
    print(f"{k[:34]:34} {a['um_rec']:>7} {a['um_pre']:>7} {a['costo']:>11,.0f} {a['impacto']:>14,.0f}")
print('─'*88)
print(f"{'TOTAL impacto mismatch':>60} {tot:>14,.0f}")
print(f"\nTotal ingredientes con mismatch: {len(agg)}")
