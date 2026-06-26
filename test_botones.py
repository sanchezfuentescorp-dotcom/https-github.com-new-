import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys, json
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
REST=f'{URL}/rest/v1'
H={'apikey':KEY,'Authorization':f'Bearer {KEY}','Content-Type':'application/json'}

def fid(tabla):
    r=requests.get(f'{REST}/{tabla}?select=id&limit=1',headers=H)
    d=r.json()
    return d[0]['id'] if isinstance(d,list) and d else None

# IDs válidos para FK
tall=fid('taller_tipo'); cat=fid('categoria'); prov=fid('proveedor')
ing=fid('ingrediente'); sem=fid('semana'); mod=fid('modulo')
try: cc=fid('centro_costo')
except: cc=None

# (tabla, datos del frontend, cómo borrar)
tests=[
  ('categoria', {'nombre':'__TEST__','color':'#f59e0b'}),
  ('proveedor', {'nombre':'__TEST__','numero':9999,'rut':'1-9','direccion':'x','tipo':'x','sede':'x','contacto':'x','condicion_pago':'x','correo':'x@x.cl','telefono':'9','notas':'x'}),
  ('taller_tipo', {'codigo':'__TST__','nombre':'__TEST__','descripcion':'x','color':'#3b82f6'}),
  ('semana', {'numero':999,'anio':2026,'nombre':'__TEST__','fecha_inicio':'2026-01-01','fecha_fin':'2026-01-05'}),
  ('ingrediente', {'nombre':'__TEST__','codigo':None,'costo_neto':100,'unidad':'kg','formato_envasado':'x','stock_minimo':0,'categoria_id':cat,'proveedor_id':prov,'fecha_actualizacion':'2026-01-01'}),
  ('sesion', {'cod_clase':'__TEST__','taller_tipo_id':tall,'semana_id':sem,'seccion':1,'num_alumnos':1,'fecha':'2026-01-01','centro_costo_id':cc}),
  ('inventario', {'ingrediente_id':ing,'cantidad':1,'mes':'MARZO','anio':2026,'updated_at':'2026-01-01T00:00:00Z'}),
  ('carrera', {'nombre':'__TEST__','duracion_horas':400,'overhead_porcentaje':20,'semestre':1,'descripcion':'x'}),
  ('modulo', {'nombre':'__TEST__','horas':80,'descripcion':'x'}),
  ('modulo_taller', {'modulo_id':mod,'taller_tipo_id':tall,'num_sesiones':1}),
  ('centro_costo', {'codigo':'__TST__','nombre':'__TEST__','descripcion':'x'}),
  ('presupuesto', {'taller_tipo_id':tall,'monto_presupuestado':1000,'num_alumnos_esperados':10,'periodo':'2026','notas':'x'}),
  ('precio_historial', {'ingrediente_id':ing,'precio_anterior':100,'precio_nuevo':110,'fecha':'2026-01-01','motivo':'x'}),
]

print('PRUEBA FUNCIONAL DE BOTONES "GUARDAR" (POST real + limpieza)')
print('─'*68)
print(f"{'TABLA / BOTÓN':22} {'RESULTADO'}")
print('─'*68)
ok=bad=0
for tabla,datos in tests:
    # saltar si falta FK requerida
    r=requests.post(f'{REST}/{tabla}',headers={**H,'Prefer':'return=representation'},data=json.dumps(datos))
    if r.status_code in (200,201):
        ok+=1
        rec=r.json()
        print(f'  ✓ {tabla:20} OK (crea correctamente)')
        # borrar
        rid=rec[0].get('id') if isinstance(rec,list) and rec else None
        if rid is not None:
            requests.delete(f'{REST}/{tabla}?id=eq.{rid}',headers=H)
        elif tabla=='modulo_taller':
            requests.delete(f'{REST}/{tabla}?modulo_id=eq.{mod}&taller_tipo_id=eq.{tall}&num_sesiones=eq.1',headers=H)
    else:
        bad+=1
        msg=''
        try: msg=r.json().get('message','')+' | '+(r.json().get('details') or '')
        except: msg=r.text[:120]
        print(f'  ✗ {tabla:20} ERROR {r.status_code}: {msg[:90]}')
print('─'*68)
print(f'Resultado: {ok} OK, {bad} con error')
print(f'\nFK disponibles: taller={tall} cat={cat} prov={prov} ing={ing} sem={sem} mod={mod} cc={cc}')
