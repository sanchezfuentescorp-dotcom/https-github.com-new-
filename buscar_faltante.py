import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys, collections
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}
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
# buscar leche condensada avena
r=requests.get(f"{URL}/rest/v1/ingrediente?nombre=ilike.*AVENA*&select=id,nombre,costo_neto,unidad",headers=H)
print('AVENA:', r.json())
r=requests.get(f"{URL}/rest/v1/ingrediente?nombre=ilike.*SAMUI*&select=id,nombre,costo_neto,unidad",headers=H)
print('SAMUI:', r.json())

# Re-chequear mismatches restantes
ings={i['id']:i for i in getall('ingrediente','id,nombre,unidad,costo_neto')}
si=getall('sesion_ingrediente','ingrediente_id,unidad,cantidad_total')
agg={}
for r in si:
    ing=ings.get(r['ingrediente_id'])
    if not ing: continue
    um_rec=(r['unidad'] or '').strip().lower(); um_pre=(ing['unidad'] or '').strip().lower()
    if um_rec!=um_pre:
        a=agg.setdefault(ing['nombre'],{'um_rec':um_rec,'um_pre':um_pre,'costo':ing['costo_neto'],'imp':0})
        a['imp']+=(r['cantidad_total'] or 0)*(ing['costo_neto'] or 0)
print(f'\nMISMATCHES RESTANTES: {len(agg)}')
for k,a in sorted(agg.items(),key=lambda x:-x[1]['imp']):
    print(f"  {k[:38]:40} {a['um_rec']:>4}≠{a['um_pre']:<4} ${a['costo']:>10,.1f}  impacto ${a['imp']:>13,.0f}")
