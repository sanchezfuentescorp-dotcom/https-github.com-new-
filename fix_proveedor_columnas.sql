-- ================================================================
-- FIX botón "Nuevo proveedor" — agrega columnas faltantes
-- Ejecutar UNA vez en Supabase → SQL Editor → New Query
-- ================================================================
-- El formulario de proveedor captura RUT y Dirección, pero la tabla
-- no tiene esas columnas, por lo que el botón "Guardar" fallaba con
-- "Could not find the 'direccion' column". Esto las agrega.
-- No afecta datos existentes.
-- ================================================================

ALTER TABLE proveedor ADD COLUMN IF NOT EXISTS rut       TEXT;
ALTER TABLE proveedor ADD COLUMN IF NOT EXISTS direccion TEXT;

-- Verificación
SELECT column_name FROM information_schema.columns
WHERE table_name = 'proveedor' ORDER BY ordinal_position;
