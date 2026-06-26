import sys
import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
sys.stdout.reconfigure(encoding='utf-8')
import requests
import json
import os

PAT      = os.getenv('SUPABASE_PAT')
REF      = 'qptywalaronkaaxgcwfw'
HEADERS  = {'Authorization': f'Bearer {PAT}', 'Content-Type': 'application/json'}
API_URL  = f'https://api.supabase.com/v1/projects/{REF}/database/query'

SQL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SETUP_COMPLETO.sql')

def run_query(sql, label=''):
    r = requests.post(API_URL, headers=HEADERS, json={'query': sql})
    ok = r.status_code in (200, 201)
    status = 'OK' if ok else f'ERROR {r.status_code}'
    msg = '' if ok else f': {r.text[:200]}'
    print(f'  [{status}] {label}{msg}')
    return ok

def split_statements(sql):
    # Split on semicolons but respect $$ blocks (DO $$ ... $$)
    statements = []
    current = []
    in_dollar = False
    for line in sql.splitlines():
        stripped = line.strip()
        if stripped.startswith('--') or not stripped:
            continue
        if '$$' in line:
            in_dollar = not in_dollar
        current.append(line)
        if not in_dollar and stripped.endswith(';'):
            stmt = '\n'.join(current).strip()
            if stmt and stmt != ';':
                statements.append(stmt)
            current = []
    if current:
        stmt = '\n'.join(current).strip()
        if stmt:
            statements.append(stmt)
    return statements

print('=' * 52)
print(' Ejecutando SETUP_COMPLETO.sql en Supabase')
print('=' * 52)

with open(SQL_FILE, 'r', encoding='utf-8') as f:
    sql_content = f.read()

# Intentar ejecutar todo de una vez primero
print('Intentando ejecucion completa...')
r = requests.post(API_URL, headers=HEADERS, json={'query': sql_content})
if r.status_code in (200, 201):
    print('  [OK] Todo el schema ejecutado exitosamente')
    print()
    print('Setup completo. La base de datos esta lista.')
    print('=' * 52)
    sys.exit(0)

print(f'  Ejecucion completa no disponible ({r.status_code}), ejecutando por bloques...')
print()

# Ejecutar por bloques separados
statements = split_statements(sql_content)
print(f'Total de sentencias: {len(statements)}')
print()

errors = 0
for i, stmt in enumerate(statements, 1):
    # Tomar las primeras palabras como label
    label = ' '.join(stmt.split()[:5]).replace('\n', ' ')[:60]
    ok = run_query(stmt, label)
    if not ok:
        errors += 1

print()
if errors == 0:
    print('Todas las sentencias ejecutadas correctamente.')
else:
    print(f'Completado con {errors} error(es). Ver detalles arriba.')
print('=' * 52)
