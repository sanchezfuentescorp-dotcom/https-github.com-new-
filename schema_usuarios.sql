-- ══════════════════════════════════════════════════════
-- CULINARY · Sistema de usuarios — ejecutar en SQL Editor
-- ══════════════════════════════════════════════════════

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
