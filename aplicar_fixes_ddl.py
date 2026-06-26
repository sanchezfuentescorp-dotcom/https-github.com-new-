import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
PAT=os.getenv('SUPABASE_PAT'); REF='qptywalaronkaaxgcwfw'
H={'Authorization':f'Bearer {PAT}','Content-Type':'application/json'}
API=f'https://api.supabase.com/v1/projects/{REF}/database/query'

def run(sql,label):
    r=requests.post(API,headers=H,json={'query':sql})
    ok=r.status_code in (200,201)
    print(f'  [{"OK" if ok else "ERROR "+str(r.status_code)}] {label}'+('' if ok else f': {r.text[:200]}'))
    return ok,(r.json() if ok and r.text else None)

print('═'*60)
print('APLICANDO FIXES DDL VÍA MANAGEMENT API')
print('═'*60)

# ── 1. Columnas faltantes en proveedor ──
print('\n[1] proveedor: agregar rut + direccion')
run("ALTER TABLE proveedor ADD COLUMN IF NOT EXISTS rut TEXT;","ADD rut")
run("ALTER TABLE proveedor ADD COLUMN IF NOT EXISTS direccion TEXT;","ADD direccion")

# ── 2. Recrear v_costo_taller sin fan-out ──
print('\n[2] v_costo_taller: corregir fan-out del JOIN')
run("""
CREATE OR REPLACE VIEW v_costo_taller AS
SELECT tt.id, tt.codigo, tt.nombre, tt.color,
    COALESCE(s.num_sesiones,0) AS num_sesiones,
    COALESCE(s.total_alumnos,0) AS total_alumnos,
    COALESCE(c.costo_total,0)::DECIMAL(10,2) AS costo_total,
    CASE WHEN COALESCE(s.total_alumnos,0)>0
        THEN (COALESCE(c.costo_total,0)/s.total_alumnos)::DECIMAL(10,2) ELSE 0 END AS costo_promedio_alumno
FROM taller_tipo tt
LEFT JOIN (SELECT taller_tipo_id, COUNT(*) AS num_sesiones, SUM(num_alumnos) AS total_alumnos
           FROM sesion GROUP BY taller_tipo_id) s ON s.taller_tipo_id=tt.id
LEFT JOIN (SELECT se.taller_tipo_id, SUM(si.cantidad_total*i.costo_neto) AS costo_total
           FROM sesion se JOIN sesion_ingrediente si ON si.sesion_id=se.id
           JOIN ingrediente i ON i.id=si.ingrediente_id GROUP BY se.taller_tipo_id) c ON c.taller_tipo_id=tt.id;
""","CREATE OR REPLACE VIEW v_costo_taller")

# ── VERIFICACIÓN ──
print('\n[3] VERIFICACIÓN')
ok,cols=run("SELECT column_name FROM information_schema.columns WHERE table_name='proveedor' AND column_name IN ('rut','direccion') ORDER BY column_name;","columnas proveedor")
if cols: print('     →', [c['column_name'] for c in cols])
ok,vt=run("SELECT codigo, total_alumnos, costo_promedio_alumno FROM v_costo_taller WHERE codigo='TTPIC';","v_costo_taller TTPIC")
if vt: print('     → TTPIC:', vt)
print('\n'+'═'*60)
