-- ================================================================
-- FIX v_costo_taller — corrige el "fan-out" del JOIN
-- OPCIONAL: el dashboard ya calcula el costo/alumno por taller
-- desde v_costo_sesion (correcto), así que NO depende de esta vista.
-- Ejecuta esto solo si quieres que la vista v_costo_taller en sí
-- devuelva valores correctos (p.ej. para consultas directas o
-- v_costo_modulo/carrera). SQL Editor → New Query.
-- ================================================================
-- PROBLEMA: la vista unía sesion × sesion_ingrediente y luego hacía
-- SUM(num_alumnos), contando los alumnos UNA VEZ POR CADA INGREDIENTE
-- de la sesión. Eso inflaba total_alumnos ~50x y subestimaba el
-- costo_promedio_alumno. Se separan los conteos en subconsultas.
-- No cambia ningún dato; solo corrige el cálculo de la vista.
-- ================================================================

CREATE OR REPLACE VIEW v_costo_taller AS
SELECT
    tt.id,
    tt.codigo,
    tt.nombre,
    tt.color,
    COALESCE(s.num_sesiones, 0)                         AS num_sesiones,
    COALESCE(s.total_alumnos, 0)                        AS total_alumnos,
    COALESCE(c.costo_total, 0)::DECIMAL(10,2)           AS costo_total,
    CASE WHEN COALESCE(s.total_alumnos, 0) > 0
        THEN (COALESCE(c.costo_total, 0) / s.total_alumnos)::DECIMAL(10,2)
        ELSE 0
    END                                                 AS costo_promedio_alumno
FROM taller_tipo tt
LEFT JOIN (
    -- conteo de sesiones y alumnos SIN duplicar por ingrediente
    SELECT taller_tipo_id,
           COUNT(*)          AS num_sesiones,
           SUM(num_alumnos)  AS total_alumnos
    FROM sesion
    GROUP BY taller_tipo_id
) s ON s.taller_tipo_id = tt.id
LEFT JOIN (
    -- costo total de materia prima por taller
    SELECT se.taller_tipo_id,
           SUM(si.cantidad_total * i.costo_neto) AS costo_total
    FROM sesion se
    JOIN sesion_ingrediente si ON si.sesion_id = se.id
    JOIN ingrediente i         ON i.id = si.ingrediente_id
    GROUP BY se.taller_tipo_id
) c ON c.taller_tipo_id = tt.id;

-- v_costo_modulo y v_costo_carrera dependen de esta vista y se
-- recalculan solas (referencian costo_promedio_alumno).

-- Verificación: total_alumnos debe ser realista (no decenas de miles)
SELECT codigo, num_sesiones, total_alumnos, costo_total, costo_promedio_alumno
FROM v_costo_taller ORDER BY codigo;
