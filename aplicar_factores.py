import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__),".env"))
import requests, sys, collections
sys.stdout.reconfigure(encoding='utf-8')
URL=os.getenv('SUPABASE_URL'); KEY=os.getenv('SUPABASE_KEY')
H={'apikey':KEY,'Authorization':f'Bearer {KEY}','Content-Type':'application/json'}
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
def clp(v): return f"${v:,.0f}".replace(',','.')

# ════════════════════════════════════════════════════════════════════════
# FACTOR = cantidad de la UNIDAD-DE-PRECIO (actual) por 1 UNIDAD-DE-RECETA.
#   nuevo_costo = costo_actual * FACTOR   (queda en $/unidad_receta)
# Pesos/volúmenes unitarios estándar de gastronomía (Chile).
# ════════════════════════════════════════════════════════════════════════
FACTOR = {
  # ── vendidos por kg, usados por unidad (kg por unidad) ──
  'COLAPEZ POR HOJA':0.002,            # hoja de gelatina ~2 g
  'ALGA NORI':0.003,                   # hoja nori ~3 g
  'PAPEL DE ARROZ':0.010,              # hoja ~10 g
  'CHALOTA':0.030,                     # chalota mediana ~30 g
  'MANTECA DE CACAO COLOR BLANCO':0.010,
  'PAN MOLDE BLANCO LAMINA':0.025,     # rebanada ~25 g
  'MIX MICROGREENS':0.010,
  'AJI CACHO CABRA AHUMADO':0.005,     # ají seco ~5 g
  'GALLETAS DE CHAMPAÑA':0.005,
  'GALLETAS LOTUS':0.006,
  'GALLETA DIET GULLON':0.006,
  'GALLETA MARAVILLA MCKAY':0.006,
  'MERLUZA FRESCA ENTERA':0.8,         # pescado entero ~800 g
  'CONGRIO COLORADO ENTERO FRESCO':1.0,
  'CORVINA ENTERA FRESCA UN':1.0,
  'PULPO ENTERO':2.0,
  'PAN MARRAQUETA':0.1,
  'PAN MARRAQUETA MINI':0.04,
  'AJI ROCOTO':0.04,
  'AJI CACHO DE CABRA FRESCO':0.015,
  'VIENESA':0.05,
  'FONDOS DE ALCACHOFA':0.03,
  'CERDO CHULETA CENTRO TRES CM':0.2,
  'CERDO PATAS':0.4,
  'LECHE EVAPORADA':0.41,              # lata 410 g
  'LECHE CONDENSADA DE AVENA (SAMUI)':0.4,
  'GARBANZOS EN CONSERVA WASIL 380 GR':0.38,
  'GARBANZOS COCIDOS EN CONSERVA':0.38,
  'MASA FILO':0.32,
  'TINTA DE SEPIA O CALAMAR':0.010,
  'AZAFRAN':0.0002,                    # pizca ~0,2 g
  'HUEVAS DE SALMON':0.05,
  'COL CHINA':0.5,
  'NABOS DAIKON':0.3,
  'MALVAVISCOS':0.005,
  'LEMON GRASS FRESCO':0.015,
  'CHUCRUT ENCURTIDO':0.2,
  'BROTES DE DECORACION':0.010,
  'MIX HOJAS VERDES':0.020,
  # ── colorantes vendidos por kg, usados por "1 un" (uso mínimo ~1 g) ──
  'COLORANTE GEL AMARILLO':0.001,'COLORANTE GEL ROJO':0.001,'COLORANTE GEL AZUL':0.001,
  'COLORANTE LIPO AMARILLO':0.001,'COLORANTE LIPO ROJO POLVO 20GR':0.001,
  'COLORANTE LIPO NARANJA POLVO 100GR':0.001,
  'COLORANTE LIPOSOLUBLE VERDE LIQUIDO':0.001,'COLORANTE LIPOSOLUBLE AZUL LIQUIDO':0.001,
  'COLORANTE LIPOSOLUBLE ROJO LIQUIDO':0.001,'COLORANTE LIPOSOLUBLE AMARILLO LIQUIDO':0.001,
  'COLORANTE HIDROSOLUBLE POLVO ROJO':0.001,'COLORANTE HIDROSOLUBLE POLVO VERDE PISTACHO':0.001,
  'COLORANTE HIDROSOLUBLE POLVO AMARILLO':0.001,'COLORANTE METALICO ORO':0.001,
  # ── líquidos vendidos por kg, usados por L (densidad ≈ 1) ──
  'CREMA DE COCO':1.0,'LECHE DE COCO':1.0,'GLICERINA ALIMENTARIA':1.26,'PAPAS':1.0,
  # ── vendidos por L, usados por kg (densidad ≈ 1) ──
  'YOGURT NATURAL':1.0,'YOGURT NATURAL LIGHT':1.0,
  # ── vendidos por L, usados por unidad (L por unidad/envase) ──
  'ENDULZANTE IANSA CERO K STEVIA 250 ML':0.25,'JUGO DE NARANJA':0.25,
  'CREMA DE COCO S/AZUCAR':0.4,'ACEITE PARA FREIR':1.0,
  # ── vendidos por unidad (botella/frasco), usados por L → (un por L) ──
  'MIRIN/LICOR DE ARROZ DULCE':1.43,'PISCO':1.43,'CERVEZA CRISTAL':2.86,
  # ── vendidos por unidad (frasco/atado), usados por kg → (un por kg) ──
  'MERMELADA DE PIÑA':2.0,'PASTA CURRY VERDE':2.5,'PIMIENTA CAYENA':2.0,
  'PERLAS DE TAPIOCA':1.0,'ENELDO FRESCO':20.0,'SALVIA FRESCA':20.0,'ESTRAGON FRESCO':20.0,
  'TE NEGRO EN HOJA':1.0,'TE MATCHA':1.0,'PALILLO AMARILLO PERUANO':1.0,
}

# unidad de receta destino por dirección (la sacamos de sesion_ingrediente)
ings={i['nombre'].upper():i for i in getall('ingrediente','id,nombre,costo_neto,unidad')}
si=getall('sesion_ingrediente','ingrediente_id,unidad')
# unidad de receta predominante por ingrediente_id
um_rec_por_id=collections.defaultdict(collections.Counter)
for r in si:
    um_rec_por_id[r['ingrediente_id']][(r['unidad'] or '').strip().lower()]+=1
def um_receta(iid):
    c=um_rec_por_id.get(iid)
    return c.most_common(1)[0][0] if c else None

print('APLICANDO FACTORES DE CONVERSIÓN PESO↔UNIDAD')
print('─'*92)
print(f"{'INGREDIENTE':37} {'$/u antes':>11} {'um_a':>4} {'factor':>7} {'$/u después':>12} {'um_d':>4}")
print('─'*92)
aplicados=0; no_encontrados=[]
for nom,factor in FACTOR.items():
    ing=ings.get(nom.upper())
    if not ing:
        no_encontrados.append(nom); continue
    um_d=um_receta(ing['id']) or ing['unidad']
    antes=ing['costo_neto'] or 0
    nuevo=round(antes*factor,4)
    r=requests.patch(f"{URL}/rest/v1/ingrediente?id=eq.{ing['id']}",headers={**H,'Prefer':'return=minimal'},
                     json={'costo_neto':nuevo,'unidad':um_d})
    ok=r.status_code in (200,204)
    flag='' if ok else ' ✗'+str(r.status_code)
    print(f"{nom[:37]:37} {antes:>11,.1f} {ing['unidad']:>4} {factor:>7} {nuevo:>12,.2f} {um_d:>4}{flag}")
    if ok: aplicados+=1

print('─'*92)
print(f'Aplicados: {aplicados}/{len(FACTOR)}')
if no_encontrados:
    print(f'No encontrados en BD ({len(no_encontrados)}): {no_encontrados}')
