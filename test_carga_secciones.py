import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}'}

# Todas las queries GET del código (sección → query)
queries=[
 ('loadAll: ingrediente','ingrediente','order=nombre&activo=eq.true&select=*,categoria(nombre,color),proveedor(nombre,sede)'),
 ('loadAll: proveedor','proveedor','order=nombre&activo=eq.true'),
 ('loadAll: categoria','categoria','order=nombre'),
 ('loadAll: semana','semana','order=numero'),
 ('loadAll: taller_tipo','taller_tipo','order=codigo'),
 ('loadAll: v_costo_taller','v_costo_taller','order=costo_promedio_alumno.desc'),
 ('loadAll: v_costo_categoria','v_costo_categoria','order=costo_total.desc'),
 ('loadAll: v_top_ingredientes','v_top_ingredientes','order=costo_total.desc&limit=10'),
 ('loadAll: v_costo_carrera','v_costo_carrera','order=costo_total.desc'),
 ('loadAll: inventario','inventario','order=updated_at.desc'),
 ('loadAll: modulo','modulo','order=nombre&select=id'),
 ('loadAll: centro_costo','centro_costo','order=nombre'),
 ('dashboard: v_costo_sesion','v_costo_sesion','order=fecha.desc'),
 ('compras: v_lista_compras','v_lista_compras','estado=eq.COMPRAR'),
 ('inventario full','inventario','order=updated_at.desc&select=*,ingrediente(nombre,unidad,categoria(nombre))'),
 ('compras full','v_lista_compras','order=estado.desc,categoria.asc,nombre.asc'),
 ('carreras: carrera_modulo','carrera_modulo','select=carrera_id,modulo_id'),
 ('carreras: modulo_taller','modulo_taller','select=modulo_id,taller_tipo_id'),
 ('modulos full','modulo','order=nombre&select=*'),
 ('modulo_taller full','modulo_taller','select=modulo_id,taller_tipo_id,num_sesiones,taller_tipo(codigo,nombre)'),
 ('semanas','semana','order=anio,numero'),
 ('historial','precio_historial','order=fecha.desc&select=*,ingrediente(nombre,unidad)'),
 ('usuarios','usuarios','order=usuario.asc'),
 ('presupuesto','presupuesto','select=*'),
 ('sesion+cc','sesion','select=id,centro_costo_id'),
]
print('TEST DE CARGA DE TODAS LAS SECCIONES (queries GET reales)')
print('─'*72)
ok=bad=0
for nombre,tabla,q in queries:
    r=requests.get(f'{URL}/rest/v1/{tabla}?{q}',headers={**H,'Range':'0-1'})
    good=r.status_code in (200,206)
    if good: ok+=1; print(f'  ✓ {nombre}')
    else:
        bad+=1
        msg=''
        try: msg=r.json().get('message','')
        except: msg=r.text[:80]
        print(f'  ✗ {nombre}  [{r.status_code}] {msg[:70]}')
print('─'*72)
print(f'{ok} OK, {bad} con error')
