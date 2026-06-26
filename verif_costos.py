import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}
def clp(v): return f"${v:,.0f}".replace(',','.')

print('COSTO PROMEDIO POR ALUMNO – POR TALLER')
print('─'*70)
r=requests.get(f'{URL}/rest/v1/v_costo_taller?order=costo_promedio_alumno.desc',headers=H)
for t in r.json():
    print(f"  {t['codigo']:7} {t['nombre'][:38]:40} {clp(t['costo_promedio_alumno']):>12} /alu  ({t['num_sesiones']} ses)")

print('\nTOP 10 INGREDIENTES POR COSTO ACUMULADO')
print('─'*70)
r=requests.get(f'{URL}/rest/v1/v_top_ingredientes?order=costo_total.desc&limit=10',headers=H)
data=r.json()
if isinstance(data,list):
    for t in data:
        print(f"  {t['nombre'][:38]:40} {clp(t['costo_total']):>14}")
else:
    print('  (respuesta:',data,')')

print('\nMUESTRA: 5 SESIONES CON COSTO')
print('─'*70)
r=requests.get(f'{URL}/rest/v1/v_costo_sesion?order=costo_total.desc&limit=5',headers=H)
for s in r.json():
    print(f"  {s['cod_clase']:22} {s['taller_codigo']:6} alu:{s['num_alumnos']:3} total:{clp(s['costo_total']):>12} /alu:{clp(s['costo_por_alumno']):>10}")

print('\nINGREDIENTES CON PRECIO $0 (requieren actualización)')
print('─'*70)
r=requests.get(f'{URL}/rest/v1/ingrediente?costo_neto=eq.0&select=nombre,unidad&order=nombre',headers=H)
for i in r.json(): print(f"  - {i['nombre']} ({i['unidad']})")
