import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}

# Buscar la sesion TPESES13-SEP-S4
r=requests.get(f"{URL}/rest/v1/sesion?cod_clase=eq.TPESES13-SEP-S4&select=id,num_alumnos",headers=H)
sid=r.json()[0]['id']; n=r.json()[0]['num_alumnos']
print(f'Sesion TPESES13-SEP-S4  id={sid}  alumnos={n}\n')

r=requests.get(f"{URL}/rest/v1/sesion_ingrediente?sesion_id=eq.{sid}&select=cantidad_unit,cantidad_total,unidad,ingrediente(nombre,costo_neto,unidad)",headers=H)
print(f"{'INGREDIENTE':35} {'cant_u':>8} {'cant_tot':>9} {'um_rec':>6} {'$/u':>10} {'um_ing':>6} {'SUBTOTAL':>12}")
print('─'*95)
tot=0
filas=[]
for si in r.json():
    ing=si['ingrediente']
    sub=si['cantidad_total']*ing['costo_neto']
    tot+=sub
    filas.append((ing['nombre'],si['cantidad_unit'],si['cantidad_total'],si['unidad'],ing['costo_neto'],ing['unidad'],sub))
for f in sorted(filas,key=lambda x:-x[6])[:20]:
    print(f"{f[0][:35]:35} {f[1]:>8} {f[2]:>9} {f[3]:>6} {f[4]:>10} {f[5]:>6} {f[6]:>12,.0f}")
print('─'*95)
print(f"{'TOTAL':>90} {tot:>12,.0f}")
