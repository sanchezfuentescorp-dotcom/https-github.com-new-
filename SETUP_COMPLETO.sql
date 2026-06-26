-- ══════════════════════════════════════════════════════════════════════
-- CULINARY · SISTEMA DE COSTOS — Setup completo
-- Ejecutar UNA VEZ en Supabase SQL Editor (Dashboard → SQL Editor → New query)
-- ══════════════════════════════════════════════════════════════════════
-- Incluye: tablas, vistas, datos de muestra, usuarios y extras
-- ══════════════════════════════════════════════════════════════════════

-- ────────────────────────────────────────────────────────────────────
-- ESQUEMA PRINCIPAL — Tablas, vistas, datos de muestra, RLS
-- ────────────────────────────────────────────────────────────────────

CREATE TABLE categoria (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL UNIQUE,
    color VARCHAR(7) DEFAULT '#6b7280'
);

CREATE TABLE proveedor (
    id SERIAL PRIMARY KEY,
    numero INTEGER,
    nombre VARCHAR(200) NOT NULL,
    tipo VARCHAR(150),
    sede VARCHAR(100),
    condicion_pago VARCHAR(50) DEFAULT '30 dias',
    contacto VARCHAR(200),
    correo TEXT,
    telefono VARCHAR(150),
    banco VARCHAR(100),
    cuenta VARCHAR(50),
    notas TEXT,
    activo BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE ingrediente (
    id SERIAL PRIMARY KEY,
    codigo VARCHAR(20) UNIQUE,
    nombre VARCHAR(200) NOT NULL,
    costo_neto DECIMAL(10,2) DEFAULT 0 CHECK (costo_neto >= 0),
    unidad VARCHAR(20) DEFAULT 'kg',
    categoria_id INTEGER REFERENCES categoria(id),
    proveedor_id INTEGER REFERENCES proveedor(id),
    activo BOOLEAN DEFAULT TRUE,
    fecha_actualizacion DATE DEFAULT CURRENT_DATE,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE taller_tipo (
    id SERIAL PRIMARY KEY,
    codigo VARCHAR(20) NOT NULL UNIQUE,
    nombre VARCHAR(150) NOT NULL,
    descripcion TEXT,
    color VARCHAR(7) DEFAULT '#3b82f6'
);

CREATE TABLE semana (
    id SERIAL PRIMARY KEY,
    numero INTEGER NOT NULL,
    anio INTEGER NOT NULL DEFAULT 2025,
    nombre VARCHAR(100),
    fecha_inicio DATE,
    fecha_fin DATE,
    UNIQUE(numero, anio)
);

CREATE TABLE sesion (
    id SERIAL PRIMARY KEY,
    cod_clase VARCHAR(80),
    taller_tipo_id INTEGER REFERENCES taller_tipo(id),
    semana_id INTEGER REFERENCES semana(id),
    fecha DATE,
    seccion INTEGER DEFAULT 1,
    num_alumnos INTEGER DEFAULT 0 CHECK (num_alumnos >= 0),
    profesor VARCHAR(150),
    notas TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE sesion_ingrediente (
    id SERIAL PRIMARY KEY,
    sesion_id INTEGER NOT NULL REFERENCES sesion(id) ON DELETE CASCADE,
    ingrediente_id INTEGER NOT NULL REFERENCES ingrediente(id),
    cantidad_unit DECIMAL(12,5) DEFAULT 0,
    cantidad_total DECIMAL(12,5) DEFAULT 0,
    unidad VARCHAR(20),
    UNIQUE(sesion_id, ingrediente_id)
);

CREATE TABLE inventario (
    id SERIAL PRIMARY KEY,
    ingrediente_id INTEGER NOT NULL REFERENCES ingrediente(id),
    cantidad DECIMAL(12,4) DEFAULT 0,
    mes VARCHAR(20),
    anio INTEGER DEFAULT 2025,
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(ingrediente_id, mes, anio)
);

-- ============================================================
-- ESTRUCTURA ACADÉMICA (Carreras → Módulos → Talleres)
-- ============================================================

CREATE TABLE modulo (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(150) NOT NULL,
    descripcion TEXT,
    horas INTEGER DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE modulo_taller (
    modulo_id INTEGER NOT NULL REFERENCES modulo(id) ON DELETE CASCADE,
    taller_tipo_id INTEGER NOT NULL REFERENCES taller_tipo(id),
    num_sesiones INTEGER DEFAULT 1,
    PRIMARY KEY (modulo_id, taller_tipo_id)
);

CREATE TABLE carrera (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(150) NOT NULL,
    descripcion TEXT,
    duracion_horas INTEGER DEFAULT 0,
    overhead_porcentaje DECIMAL(5,2) DEFAULT 20,
    activa BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE carrera_modulo (
    carrera_id INTEGER NOT NULL REFERENCES carrera(id) ON DELETE CASCADE,
    modulo_id INTEGER NOT NULL REFERENCES modulo(id),
    orden INTEGER DEFAULT 1,
    PRIMARY KEY (carrera_id, modulo_id)
);

-- ============================================================
-- VISTAS DE COSTOS
-- ============================================================

-- Costo real por sesión de clase
CREATE VIEW v_costo_sesion AS
SELECT
    s.id,
    s.cod_clase,
    tt.codigo   AS taller_codigo,
    tt.nombre   AS taller_nombre,
    tt.color    AS taller_color,
    sem.numero  AS semana_numero,
    sem.anio    AS semana_anio,
    s.fecha,
    s.seccion,
    s.num_alumnos,
    s.profesor,
    COALESCE(SUM(si.cantidad_total * i.costo_neto), 0)::DECIMAL(10,2) AS costo_total,
    CASE WHEN s.num_alumnos > 0
        THEN (COALESCE(SUM(si.cantidad_total * i.costo_neto), 0) / s.num_alumnos)::DECIMAL(10,2)
        ELSE 0
    END AS costo_por_alumno
FROM sesion s
LEFT JOIN taller_tipo tt ON tt.id = s.taller_tipo_id
LEFT JOIN semana sem      ON sem.id = s.semana_id
LEFT JOIN sesion_ingrediente si ON si.sesion_id = s.id
LEFT JOIN ingrediente i   ON i.id = si.ingrediente_id
GROUP BY s.id, s.cod_clase, tt.codigo, tt.nombre, tt.color,
         sem.numero, sem.anio, s.fecha, s.seccion, s.num_alumnos, s.profesor;

-- Costo agregado por tipo de taller
CREATE VIEW v_costo_taller AS
SELECT
    tt.id,
    tt.codigo,
    tt.nombre,
    tt.color,
    COUNT(DISTINCT s.id)                                              AS num_sesiones,
    COALESCE(SUM(s.num_alumnos), 0)                                  AS total_alumnos,
    COALESCE(SUM(si.cantidad_total * i.costo_neto), 0)::DECIMAL(10,2) AS costo_total,
    CASE WHEN COALESCE(SUM(s.num_alumnos), 0) > 0
        THEN (COALESCE(SUM(si.cantidad_total * i.costo_neto), 0) / SUM(s.num_alumnos))::DECIMAL(10,2)
        ELSE 0
    END AS costo_promedio_alumno
FROM taller_tipo tt
LEFT JOIN sesion s             ON s.taller_tipo_id = tt.id
LEFT JOIN sesion_ingrediente si ON si.sesion_id = s.id
LEFT JOIN ingrediente i        ON i.id = si.ingrediente_id
GROUP BY tt.id, tt.codigo, tt.nombre, tt.color;

-- Distribución de costos por categoría
CREATE VIEW v_costo_categoria AS
SELECT
    c.id,
    c.nombre AS categoria,
    c.color,
    COALESCE(SUM(si.cantidad_total * i.costo_neto), 0)::DECIMAL(10,2) AS costo_total,
    COUNT(DISTINCT i.id) AS num_ingredientes
FROM categoria c
LEFT JOIN ingrediente i        ON i.categoria_id = c.id
LEFT JOIN sesion_ingrediente si ON si.ingrediente_id = i.id
GROUP BY c.id, c.nombre, c.color;

-- Top ingredientes por costo total acumulado
CREATE VIEW v_top_ingredientes AS
SELECT
    i.id,
    i.nombre,
    i.codigo,
    i.costo_neto,
    i.unidad,
    c.nombre AS categoria,
    COALESCE(SUM(si.cantidad_total), 0)::DECIMAL(12,4) AS cantidad_usada,
    COALESCE(SUM(si.cantidad_total * i.costo_neto), 0)::DECIMAL(10,2) AS costo_total
FROM ingrediente i
LEFT JOIN categoria c          ON c.id = i.categoria_id
LEFT JOIN sesion_ingrediente si ON si.ingrediente_id = i.id
GROUP BY i.id, i.nombre, i.codigo, i.costo_neto, i.unidad, c.nombre
ORDER BY costo_total DESC;

-- Lista de compras (necesario vs stock)
CREATE VIEW v_lista_compras AS
SELECT
    i.id,
    i.codigo,
    i.nombre,
    i.unidad,
    i.costo_neto,
    c.nombre AS categoria,
    p.nombre AS proveedor,
    COALESCE(SUM(si.cantidad_total), 0)::DECIMAL(12,4) AS cantidad_necesaria,
    COALESCE(inv.cantidad, 0)::DECIMAL(12,4)            AS stock_actual,
    GREATEST(0, COALESCE(SUM(si.cantidad_total), 0) - COALESCE(inv.cantidad, 0))::DECIMAL(12,4) AS cantidad_comprar,
    CASE WHEN COALESCE(SUM(si.cantidad_total), 0) > COALESCE(inv.cantidad, 0)
        THEN 'COMPRAR' ELSE 'OK'
    END AS estado
FROM ingrediente i
LEFT JOIN categoria c  ON c.id = i.categoria_id
LEFT JOIN proveedor p  ON p.id = i.proveedor_id
LEFT JOIN sesion_ingrediente si ON si.ingrediente_id = i.id
LEFT JOIN (
    SELECT ingrediente_id, SUM(cantidad) AS cantidad
    FROM inventario
    GROUP BY ingrediente_id
) inv ON inv.ingrediente_id = i.id
GROUP BY i.id, i.codigo, i.nombre, i.unidad, i.costo_neto,
         c.nombre, p.nombre, inv.cantidad;

-- Costo por módulo (usando sesiones reales)
CREATE VIEW v_costo_modulo AS
SELECT
    m.id,
    m.nombre,
    m.horas,
    COALESCE(
        SUM(vct.costo_promedio_alumno * mt.num_sesiones), 0
    )::DECIMAL(10,2) AS costo_modulo_por_alumno,
    COUNT(mt.taller_tipo_id) AS num_talleres
FROM modulo m
LEFT JOIN modulo_taller mt ON mt.modulo_id = m.id
LEFT JOIN v_costo_taller vct ON vct.id = mt.taller_tipo_id
GROUP BY m.id, m.nombre, m.horas;

-- Costo por carrera (con overhead)
CREATE VIEW v_costo_carrera AS
SELECT
    ca.id,
    ca.nombre,
    ca.duracion_horas,
    ca.overhead_porcentaje,
    COALESCE(SUM(vcm.costo_modulo_por_alumno), 0)::DECIMAL(10,2) AS costo_materia_prima,
    (COALESCE(SUM(vcm.costo_modulo_por_alumno), 0) * (1 + ca.overhead_porcentaje / 100))::DECIMAL(10,2) AS costo_total,
    COUNT(cm.modulo_id) AS num_modulos
FROM carrera ca
LEFT JOIN carrera_modulo cm ON cm.carrera_id = ca.id
LEFT JOIN v_costo_modulo vcm ON vcm.id = cm.modulo_id
GROUP BY ca.id, ca.nombre, ca.duracion_horas, ca.overhead_porcentaje;

-- ============================================================
-- ROW LEVEL SECURITY
-- ============================================================

DO $$ DECLARE t text;
BEGIN
  FOR t IN SELECT unnest(ARRAY['categoria','proveedor','ingrediente','taller_tipo',
    'semana','sesion','sesion_ingrediente','inventario',
    'modulo','modulo_taller','carrera','carrera_modulo'])
  LOOP
    EXECUTE format('ALTER TABLE %I ENABLE ROW LEVEL SECURITY', t);
    EXECUTE format('CREATE POLICY "public_all" ON %I FOR ALL TO anon USING (true) WITH CHECK (true)', t);
  END LOOP;
END $$;

GRANT SELECT ON v_costo_sesion, v_costo_taller, v_costo_categoria,
               v_top_ingredientes, v_lista_compras, v_costo_modulo, v_costo_carrera TO anon;

-- ============================================================
-- DATOS BASE (categorías y talleres reales de Culinary)
-- ============================================================

INSERT INTO categoria (nombre, color) VALUES
('ABARROTE',        '#f59e0b'),
('FRUTA VERDURA',   '#10b981'),
('CARNES',          '#ef4444'),
('PESCADO MARISCO', '#3b82f6'),
('LACTEO',          '#8b5cf6'),
('HUEVO',           '#f97316'),
('CHOCOLATERIA',    '#92400e'),
('CONGELADOS',      '#06b6d4'),
('LICOR',           '#6366f1'),
('AGUA',            '#0ea5e9'),
('QUIMICOS',        '#64748b'),
('OTROS',           '#9ca3af');

INSERT INTO taller_tipo (codigo, nombre, descripcion, color) VALUES
('TC1',  'Técnicas de Cocina 1',    'Fundamentos culinarios y técnicas básicas', '#ef4444'),
('TC2',  'Técnicas de Cocina 2',    'Técnicas intermedias y cocina regional',    '#dc2626'),
('TP1',  'Pastelería 1',            'Pastelería clásica, masas y cremas',        '#f59e0b'),
('TP2',  'Pastelería 2',            'Pastelería avanzada, chocolate y decoración','#d97706'),
('TDLC', 'Taller De La Cocina',     'Taller avanzado cocina de autor',           '#8b5cf6'),
('TIC',  'Introducción a la Cocina','Primer acercamiento culinario',             '#10b981'),
('TM',   'Taller Mediterráneo',     'Cocina mediterránea y del mundo',           '#3b82f6'),
('TPP',  'Panadería y Pizzería',    'Pan artesanal, masas fermentadas y pizzas', '#f97316');

INSERT INTO proveedor (numero, nombre, tipo, sede, condicion_pago, contacto, correo, telefono, notas) VALUES
(4,  'Bidfood Chile SA',                'ABARROTES',          'SANTIAGO',  '30 dias', 'ARELI',              'pagos@bidfood.cl',         '56225994444', 'Verificar por plataforma'),
(5,  'Caserita en Casa SPA',            'FRUTAS Y VERDURAS',  'SANTIAGO',  '30 dias', 'Francisco Montaner', 'caseritaencasa@gmail.com', '56952069703', 'Listado de precio'),
(20, 'Gourmet Select LTDA',             'ESPECIAS',           'STGO/VIÑA', '30 dias', 'Génesis Sanchez',    'chef@gourmetselect.cl',    '56944921093', 'Mandar cotizacion de Productos'),
(23, 'La Vinoteca',                     'VINOS Y LICORES',    'STGO/VIÑA', '30 dias', 'Johana Quiroga',     'andrea.aristizabal@lavinoteca.cl','56228292221','Cotizar vinos de maridaje'),
(34, 'Soc. Ganadera Rodriguez y Grassi','CARNES',             'VIÑA',      '30 dias', 'Patricia Tapia',     'gr.ganaderareal@gmail.com','56983670527', 'Buen proveedor'),
(36, 'Neucober S.A',                    'CHOCOLATE',          'STGO/VIÑA', '30 dias', 'Daniel González',    'danielgonzalez@neucober.cl','56978895811','Cotizar podemos cambiar por gourmet'),
(31, 'Palais Gourmet LTDA',             'CARNES Y LACTEOS',   'SANTIAGO',  '30 dias', 'Gonzalo Rivas',      'grivas@palaisgourmet.cl',  '56942132129', 'Cotizar con Gonzalo'),
(37, 'Leche Fresca SpA',                'LECHE FRESCA',       'SANTIAGO',  '30 dias', 'Luis Infante',       'luis@frescarebeca.cl',     '56985008017', 'Cotizar productos'),
(17, 'Comercializadora Benbo Spa',      'HUEVOS',             '',          '30 dias', 'Nicolas Borobio',    'equipo@benbo.cl',          '',            'Pedir cotizacion');

INSERT INTO ingrediente (codigo, nombre, costo_neto, unidad, categoria_id, proveedor_id) VALUES
('10030', 'ACEITE DE MARAVILLA',      2304,   'L',  (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=4)),
('10040', 'ACEITE DE OLIVA',          14516,  'L',  (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=4)),
('18720', 'ACEITE PARA FREIR',        2270,   'L',  (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=4)),
('10060', 'ACEITE VEGETAL',           2300,   'L',  (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=4)),
('10570', 'AZUCAR BLANCA GRANULADA',  780,    'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=5)),
('10560', 'AZUCAR FLOR',              950,    'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=5)),
('10100', 'AGAR-AGAR',                81000,  'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=20)),
('15640', 'CREMA DE LECHE',           2800,   'L',  (SELECT id FROM categoria WHERE nombre='LACTEO'),   (SELECT id FROM proveedor WHERE numero=8)),
('15700', 'LECHE ENTERA',             980,    'L',  (SELECT id FROM categoria WHERE nombre='LACTEO'),   (SELECT id FROM proveedor WHERE numero=8)),
('15600', 'HUEVO EXTRA BLANCO',       180,    'UN', (SELECT id FROM categoria WHERE nombre='HUEVO'),    (SELECT id FROM proveedor WHERE numero=17)),
('18430', 'ZANAHORIA',                890,    'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('23450', 'CEBOLLA BLANCA',           650,    'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('17640', 'AJO CABEZA',               2200,   'UN', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('17890', 'CILANTRO',                 1500,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('18200', 'PAPAS',                    700,    'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('14910', 'LIMON DE JUGO',            1200,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('18390', 'TOMATE LARGA VIDA',        1100,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('14850', 'ARANDANOS FRESCOS',        8500,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('14880', 'FRAMBUESA FRESCA',         9200,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('14890', 'FRUTILLA FRESCA',          4800,   'kg', (SELECT id FROM categoria WHERE nombre='FRUTA VERDURA'), (SELECT id FROM proveedor WHERE numero=5)),
('14290', 'CERDO COSTILLAR',          5500,   'kg', (SELECT id FROM categoria WHERE nombre='CARNES'),    (SELECT id FROM proveedor WHERE numero=34)),
('13160', 'SAL DE MESA LOBOS',        320,    'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=5)),
('12550', 'OREGANO SECO',             4200,   'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=20)),
('10180', 'AJI EN PASTA',             3800,   'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=5)),
('12160', 'MAICENA',                  1200,   'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=4)),
('11380', 'ESENCIA DE VAINILLA',      18000,  'L',  (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=20)),
('11680', 'GLUCOSA',                  4500,   'kg', (SELECT id FROM categoria WHERE nombre='QUIMICOS'), (SELECT id FROM proveedor WHERE numero=20)),
('14540', 'COBERTURA AMARGA',         9800,   'kg', (SELECT id FROM categoria WHERE nombre='CHOCOLATERIA'), (SELECT id FROM proveedor WHERE numero=36)),
('10170', 'AJI DE COLOR',             5200,   'kg', (SELECT id FROM categoria WHERE nombre='ABARROTE'), (SELECT id FROM proveedor WHERE numero=20)),
('18510', 'ARVEJAS CONGELADAS',       1800,   'kg', (SELECT id FROM categoria WHERE nombre='CONGELADOS'), (SELECT id FROM proveedor WHERE numero=4));

-- Semanas de ejemplo
INSERT INTO semana (numero, anio, nombre, fecha_inicio, fecha_fin) VALUES
(1, 2025, 'Semana 1 - Inducción',     '2025-03-10', '2025-03-14'),
(2, 2025, 'Semana 2 - Técnicas Base', '2025-03-17', '2025-03-21'),
(3, 2025, 'Semana 3 - Proteínas',     '2025-03-24', '2025-03-28');

-- Sesiones de ejemplo (basadas en datos reales del Excel)
INSERT INTO sesion (cod_clase, taller_tipo_id, semana_id, fecha, seccion, num_alumnos, profesor) VALUES
('TCCMSES12', (SELECT id FROM taller_tipo WHERE codigo='TC1'), (SELECT id FROM semana WHERE numero=1 AND anio=2025), '2025-03-10', 1, 15, 'EMERSON O.'),
('TIPSES17',  (SELECT id FROM taller_tipo WHERE codigo='TP1'), (SELECT id FROM semana WHERE numero=1 AND anio=2025), '2025-03-10', 1, 17, 'FERNANDA M.'),
('TIRSES21',  (SELECT id FROM taller_tipo WHERE codigo='TDLC'),(SELECT id FROM semana WHERE numero=3 AND anio=2025), '2025-03-24', 1, 1,  'RAUL S.'),
('TCCHOSES1', (SELECT id FROM taller_tipo WHERE codigo='TP2'), (SELECT id FROM semana WHERE numero=3 AND anio=2025), '2025-03-24', 1, 14, 'GABRIELA B.');

-- Ingredientes de la sesión TC1 - Semana 1
INSERT INTO sesion_ingrediente (sesion_id, ingrediente_id, cantidad_unit, cantidad_total, unidad) VALUES
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='10030'), 0.08, 1.20, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='17640'), 0.30, 4.50, 'UN'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='23450'), 0.15, 2.25, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='18200'), 0.20, 3.00, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='14910'), 0.10, 1.50, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='18430'), 0.08, 1.20, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='14290'), 0.25, 3.75, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='15600'), 1.00, 15.00,'UN'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='13160'), 0.01, 0.15, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCMSES12'),
 (SELECT id FROM ingrediente WHERE codigo='17890'), 0.10, 1.50, 'kg');

-- Ingredientes de la sesión TP1 - Semana 1
INSERT INTO sesion_ingrediente (sesion_id, ingrediente_id, cantidad_unit, cantidad_total, unidad) VALUES
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='14850'), 0.09, 1.53, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='10570'), 0.15, 2.55, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='14540'), 0.07, 1.19, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='15640'), 0.15, 2.55, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='11380'), 0.003, 0.051,'L'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='14880'), 0.07, 1.19, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='14890'), 0.125,2.125,'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='11680'), 0.015,0.255,'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='15600'), 4.50, 76.5, 'UN'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='15700'), 0.10, 1.70, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TIPSES17'),
 (SELECT id FROM ingrediente WHERE codigo='12160'), 0.01, 0.17, 'kg');

-- Ingredientes sesión TDLC - Semana 3
INSERT INTO sesion_ingrediente (sesion_id, ingrediente_id, cantidad_unit, cantidad_total, unidad) VALUES
((SELECT id FROM sesion WHERE cod_clase='TIRSES21'),
 (SELECT id FROM ingrediente WHERE codigo='10030'), 1.50, 1.50, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TIRSES21'),
 (SELECT id FROM ingrediente WHERE codigo='10040'), 0.10, 0.10, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TIRSES21'),
 (SELECT id FROM ingrediente WHERE codigo='18720'), 1.50, 1.50, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TIRSES21'),
 (SELECT id FROM ingrediente WHERE codigo='10570'), 0.10, 0.10, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TIRSES21'),
 (SELECT id FROM ingrediente WHERE codigo='14850'), 0.20, 0.20, 'kg');

-- Ingredientes sesión TP2 - Semana 3
INSERT INTO sesion_ingrediente (sesion_id, ingrediente_id, cantidad_unit, cantidad_total, unidad) VALUES
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='14540'), 0.03, 0.42, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='10570'), 0.11, 1.54, 'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='10100'), 0.001,0.014,'kg'),
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='15640'), 0.08, 1.12, 'L'),
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='15600'), 2.00, 28.00,'UN'),
((SELECT id FROM sesion WHERE cod_clase='TCCHOSES1'),
 (SELECT id FROM ingrediente WHERE codigo='10560'), 0.06, 0.84, 'kg');

-- Inventario de ejemplo (mes MARZO 2025)
INSERT INTO inventario (ingrediente_id, cantidad, mes, anio) VALUES
((SELECT id FROM ingrediente WHERE codigo='10040'), 3.0,  'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='18720'), 5.0,  'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='10570'), 33.0, 'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='15700'), 0.0,  'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='14850'), 0.0,  'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='14290'), 1.5,  'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='15600'), 20.0, 'MARZO', 2025),
((SELECT id FROM ingrediente WHERE codigo='14540'), 0.5,  'MARZO', 2025);

-- Estructura académica de ejemplo
INSERT INTO modulo (nombre, descripcion, horas) VALUES
('Fundamentos Culinarios',        'Técnicas básicas, caldos, salsas madre', 80),
('Pastelería Clásica',            'Masas, cremas, chocolatería y decoración', 60),
('Cocina Avanzada',               'Técnicas modernas y cocina de autor', 80);

INSERT INTO modulo_taller (modulo_id, taller_tipo_id, num_sesiones) VALUES
((SELECT id FROM modulo WHERE nombre='Fundamentos Culinarios'), (SELECT id FROM taller_tipo WHERE codigo='TC1'), 8),
((SELECT id FROM modulo WHERE nombre='Fundamentos Culinarios'), (SELECT id FROM taller_tipo WHERE codigo='TIC'), 4),
((SELECT id FROM modulo WHERE nombre='Pastelería Clásica'),     (SELECT id FROM taller_tipo WHERE codigo='TP1'), 6),
((SELECT id FROM modulo WHERE nombre='Pastelería Clásica'),     (SELECT id FROM taller_tipo WHERE codigo='TP2'), 4),
((SELECT id FROM modulo WHERE nombre='Cocina Avanzada'),        (SELECT id FROM taller_tipo WHERE codigo='TDLC'),6),
((SELECT id FROM modulo WHERE nombre='Cocina Avanzada'),        (SELECT id FROM taller_tipo WHERE codigo='TM'), 4);

INSERT INTO carrera (nombre, descripcion, duracion_horas, overhead_porcentaje) VALUES
('Chef Profesional',         'Formación completa como chef de cocina',       1200, 25),
('Especialidad en Pastelería','Especialización en pastelería y chocolatería', 400,  20),
('Técnico en Cocina',        'Programa técnico de nivel inicial',            600,  20);

INSERT INTO carrera_modulo (carrera_id, modulo_id, orden) VALUES
((SELECT id FROM carrera WHERE nombre='Chef Profesional'), (SELECT id FROM modulo WHERE nombre='Fundamentos Culinarios'), 1),
((SELECT id FROM carrera WHERE nombre='Chef Profesional'), (SELECT id FROM modulo WHERE nombre='Pastelería Clásica'), 2),
((SELECT id FROM carrera WHERE nombre='Chef Profesional'), (SELECT id FROM modulo WHERE nombre='Cocina Avanzada'), 3),
((SELECT id FROM carrera WHERE nombre='Especialidad en Pastelería'), (SELECT id FROM modulo WHERE nombre='Pastelería Clásica'), 1),
((SELECT id FROM carrera WHERE nombre='Técnico en Cocina'), (SELECT id FROM modulo WHERE nombre='Fundamentos Culinarios'), 1);

-- ────────────────────────────────────────────────────────────────────
-- USUARIOS — Sistema de autenticación
-- ────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS usuarios (
  id            SERIAL PRIMARY KEY,
  usuario       TEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  rol           TEXT NOT NULL DEFAULT 'lectura' CHECK (rol IN ('admin','lectura')),
  nombre        TEXT,
  activo        BOOLEAN DEFAULT true,
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE usuarios ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public_select" ON usuarios FOR SELECT TO anon USING (activo = true);

-- Usuarios iniciales
-- admin    → contraseña: culinary2025
-- docente  → contraseña: culinary123
INSERT INTO usuarios (usuario, password_hash, rol, nombre) VALUES
('admin',   '5264b62f8b8b0be76247503cd6486a5b248e3afd9aa63e9554102ba709afd7e3', 'admin',   'Administrador'),
('docente', 'd1c30bd64f4bce6cd8940497b6fbb3229eed772882bff1bdfdec70201c23f78e', 'lectura', 'Docente')
ON CONFLICT (usuario) DO NOTHING;

-- ────────────────────────────────────────────────────────────────────
-- EXTRAS — Historial de precios y Presupuesto vs Real
-- ────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS precio_historial (
  id               SERIAL PRIMARY KEY,
  ingrediente_id   INTEGER REFERENCES ingrediente(id) ON DELETE CASCADE,
  precio_anterior  DECIMAL(12,2) NOT NULL DEFAULT 0,
  precio_nuevo     DECIMAL(12,2) NOT NULL DEFAULT 0,
  fecha            DATE DEFAULT CURRENT_DATE,
  motivo           TEXT,
  created_at       TIMESTAMPTZ DEFAULT NOW()
);

-- Presupuesto planificado por taller y período
CREATE TABLE IF NOT EXISTS presupuesto (
  id                    SERIAL PRIMARY KEY,
  taller_tipo_id        INTEGER REFERENCES taller_tipo(id) ON DELETE CASCADE,
  monto_presupuestado   DECIMAL(14,2) NOT NULL DEFAULT 0,
  num_alumnos_esperados INTEGER DEFAULT 0,
  periodo               TEXT,
  notas                 TEXT,
  created_at            TIMESTAMPTZ DEFAULT NOW()
);

-- RLS para nuevas tablas
ALTER TABLE precio_historial ENABLE ROW LEVEL SECURITY;
ALTER TABLE presupuesto      ENABLE ROW LEVEL SECURITY;

CREATE POLICY "public_all" ON precio_historial FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "public_all" ON presupuesto      FOR ALL TO anon USING (true) WITH CHECK (true);

