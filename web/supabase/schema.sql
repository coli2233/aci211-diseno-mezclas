-- ============================================================
-- ACI 211 · Esquema Supabase: historial de diseños + RLS
-- Ejecutar en Supabase Dashboard > SQL Editor (New query > Run)
-- ============================================================

-- 1) Tabla de diseños guardados
create table if not exists public.disenos (
  id             uuid primary key default gen_random_uuid(),
  usuario_id     uuid not null references auth.users (id) on delete cascade,
  descripcion    text,
  fc             numeric,
  datos_entrada  jsonb not null,          -- payload enviado a /api/disenar
  dosificacion   jsonb,                   -- lo que devolvió el script
  informe_nombre text,                    -- nombre del Excel generado
  salida         text,                    -- stdout crudo del script (opcional)
  creado_en      timestamptz not null default now()
);

create index if not exists disenos_usuario_recientes
  on public.disenos (usuario_id, creado_en desc);

-- 2) Row Level Security: cada usuario solo ve/crea/borra lo suyo
alter table public.disenos enable row level security;

drop policy if exists "dueno_lee_sus_disenos" on public.disenos;
create policy "dueno_lee_sus_disenos" on public.disenos
  for select using (auth.uid() = usuario_id);

drop policy if exists "dueno_inserta_sus_disenos" on public.disenos;
create policy "dueno_inserta_sus_disenos" on public.disenos
  for insert with check (auth.uid() = usuario_id);

drop policy if exists "dueno_actualiza_sus_disenos" on public.disenos;
create policy "dueno_actualiza_sus_disenos" on public.disenos
  for update using (auth.uid() = usuario_id)
  with check (auth.uid() = usuario_id);

drop policy if exists "dueno_borra_sus_disenos" on public.disenos;
create policy "dueno_borra_sus_disenos" on public.disenos
  for delete using (auth.uid() = usuario_id);

-- Nota: anon NO tiene ninguna política -> sin sesión no se puede leer ni escribir.

-- 3) (Opcional) bucket para archivar los Excel generados por la API
--    Descomenta si luego quieres subir el .xlsx a Supabase Storage:
--
-- insert into storage.buckets (id, name, public)
--   values ('informes-aci211', 'informes-aci211', false)
--   on conflict (id) do nothing;
--
-- create policy "dueno_lee_sus_informes" on storage.objects
--   for select using (
--     bucket_id = 'informes-aci211'
--     and (storage.foldername(name))[1] = auth.uid()::text
--   );
--
-- create policy "dueno_sube_sus_informes" on storage.objects
--   for insert with check (
--     bucket_id = 'informes-aci211'
--     and (storage.foldername(name))[1] = auth.uid()::text
--   );
