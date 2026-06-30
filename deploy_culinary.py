import sys
import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
sys.stdout.reconfigure(encoding='utf-8')
import requests
import os

TOKEN    = os.getenv('PA_TOKEN')
USERNAME = 'Culinary'
DOMAIN   = 'culinary.pythonanywhere.com'
HEADERS  = {'Authorization': f'Token {TOKEN}'}
PA_BASE  = 'https://www.pythonanywhere.com/api/v0'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FLASK_APP_CONTENT = '''\
import os
from flask import Flask, send_from_directory

app = Flask(__name__)
HOME = os.path.dirname(os.path.abspath(__file__))

@app.route('/')
def index():
    return send_from_directory(HOME, 'index.html')
'''

WSGI_CONTENT = f'''\
import sys
sys.path.insert(0, '/home/{USERNAME}')
from flask_app import app as application
'''

def upload_text(remote_path, content, label):
    print(f'  Subiendo {label}...')
    url = f'{PA_BASE}/user/{USERNAME}/files/path{remote_path}'
    r = requests.post(url, files={'content': ('file', content.encode('utf-8'))}, headers=HEADERS)
    if r.status_code in (200, 201):
        print(f'    OK - HTTP {r.status_code}')
    else:
        print(f'    ERROR - HTTP {r.status_code}: {r.text}')
        sys.exit(1)

def upload_binary(remote_path, local_path, label):
    print(f'  Subiendo {label}...')
    url = f'{PA_BASE}/user/{USERNAME}/files/path{remote_path}'
    with open(local_path, 'rb') as f:
        r = requests.post(url, files={'content': f}, headers=HEADERS)
    if r.status_code in (200, 201):
        print(f'    OK - HTTP {r.status_code}')
    else:
        print(f'    ERROR - HTTP {r.status_code}: {r.text}')
        sys.exit(1)

def ensure_webapp():
    print('Verificando web app...')
    url = f'{PA_BASE}/user/{USERNAME}/webapps/{DOMAIN}/'
    r = requests.get(url, headers=HEADERS)
    if r.status_code == 200:
        print(f'  Web app existente encontrada.')
        return
    print('  Creando nueva web app Flask...')
    r = requests.post(
        f'{PA_BASE}/user/{USERNAME}/webapps/',
        data={'domain_name': DOMAIN, 'python_version': 'python310'},
        headers=HEADERS
    )
    if r.status_code in (200, 201):
        print(f'  Web app creada - HTTP {r.status_code}')
    else:
        print(f'  ERROR creando web app - HTTP {r.status_code}: {r.text}')
        sys.exit(1)

def configure_wsgi():
    print('Configurando WSGI...')
    wsgi_path = f'/var/www/{USERNAME.lower()}_pythonanywhere_com_wsgi.py'
    upload_text(wsgi_path, WSGI_CONTENT, 'wsgi.py')

def reload_app():
    print('Recargando web app...')
    url = f'{PA_BASE}/user/{USERNAME}/webapps/{DOMAIN}/reload/'
    r = requests.post(url, headers=HEADERS)
    if r.ok:
        print('  OK - app recargada')
    else:
        print(f'  ERROR - HTTP {r.status_code}: {r.text}')

def git_commit_push():
    import subprocess, datetime
    print('Sincronizando con GitHub...')
    fecha = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    cmds = [
        ['git', '-C', BASE_DIR, 'add', '-A'],
        ['git', '-C', BASE_DIR, 'commit', '-m', f'deploy {fecha}'],
        ['git', '-C', BASE_DIR, 'push', 'origin', 'master'],
    ]
    sfcorp_token = os.getenv('SFCORP_GH_TOKEN', '')
    sfcorp_repos = [
        f'https://{sfcorp_token}@github.com/sanchezfuentescorp-dotcom/CULINARY_SYSTEM.git',
        f'https://{sfcorp_token}@github.com/sanchezfuentescorp-dotcom/https-github.com-new-.git',
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0 and 'nothing to commit' not in r.stdout + r.stderr:
            print(f'  Git: {(r.stdout or r.stderr).strip()[:120]}')
        elif 'master' in r.stdout + r.stderr:
            print(f'  OK - pushed a GitHub (origin)')
    if sfcorp_token:
        for url in sfcorp_repos:
            r = subprocess.run(['git', '-C', BASE_DIR, 'push', url, 'master'],
                               capture_output=True, text=True)
            repo_name = url.split('/')[-1].replace('.git','')
            if r.returncode != 0:
                print(f'  SF Corp [{repo_name}]: {(r.stdout or r.stderr).strip()[:120]}')
            else:
                print(f'  OK - pushed a SF Corp ({repo_name})')

if __name__ == '__main__':
    print('=' * 52)
    print(' Deploy Culinary Costos -> PythonAnywhere')
    print('=' * 52)

    ensure_webapp()

    print('Subiendo archivos...')
    upload_text(f'/home/{USERNAME}/flask_app.py', FLASK_APP_CONTENT, 'flask_app.py')
    upload_binary(f'/home/{USERNAME}/index.html',
                  os.path.join(BASE_DIR, 'index.html'), 'index.html')

    configure_wsgi()
    reload_app()

    git_commit_push()

    print()
    print(f'Listo: https://{DOMAIN}/')
    print('=' * 52)
