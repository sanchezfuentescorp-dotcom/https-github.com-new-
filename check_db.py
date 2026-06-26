import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')

URL = os.getenv('SUPABASE_URL')
KEY = os.getenv('SUPABASE_KEY')
H = {'apikey': KEY, 'Authorization': f'Bearer {KEY}'}

def count(tabla):
    r = requests.get(f'{URL}/rest/v1/{tabla}', headers={**H, 'Prefer':'count=exact','Range':'0-0'})
    if r.status_code not in (200,206):
        return f'ERROR {r.status_code}: {r.text[:120]}'
    cr = r.headers.get('content-range','?/?')
    return cr.split('/')[-1]

for t in ['categoria','proveedor','taller_tipo','ingrediente','sesion','sesion_ingrediente','semana','inventario','modulo','carrera']:
    print(f'  {t:22} → {count(t)}')
