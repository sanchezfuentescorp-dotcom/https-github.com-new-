-- ══════════════════════════════════════════════════════════════
-- CULINARY · Schema extras v2 — Historial Precios + Presupuesto
-- Ejecutar en Supabase SQL Editor DESPUÉS del schema.sql original
-- ══════════════════════════════════════════════════════════════

-- Historial de cambios de precio por ingrediente
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
