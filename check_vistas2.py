import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}
# traer TODO (sin limit) para forzar materialización completa
for v in ['v_costo_sesion','v_costo_taller','v_costo_categoria','v_top_ingredientes','v_lista_compras','v_costo_modulo','v_costo_carrera']:
    r=requests.get(f'{URL}/rest/v1/{v}?select=*',headers={**H,'Range':'0-4999'})
    d=r.json()
    ok=isinstance(d,list)
    print(f"  {'✓' if ok else '✗ OVERFLOW'} {v:20} {r.status_code} {('('+str(len(d))+' filas)') if ok else d.get('message','')[:50]}")
