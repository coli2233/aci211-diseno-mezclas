-- ============================================================
-- ACI 211 · Esquema para SUPABASE (Equipo 6)
-- Este archivo lo ejecutamos en el proyecto Supabase "aci211"
-- (uiudnidegokdnnisjpho · us-east-1) desde su SQL Editor, y ya
-- está aplicado en ese servidor. Con él creamos el historial de
-- diseños ("Mis diseños") y lo protegemos con RLS para que cada
-- usuario vea únicamente lo suyo.
-- ============================================================

-- 1) Creamos la tabla de diseños guardados (vive en Supabase)
create table if not exists public.disenos (
  id             uuid primary key default gen_random_uuid(),
  usuario_id     uuid not null references auth.users (id) on delete cascade,  -- dueño (Supabase Auth)
  descripcion    text,
  fc             numeric,
  datos_entrada  jsonb not null,          -- lo que enviamos a /api/disenar
  dosificacion   jsonb,                   -- lo que devolvió el script
  informe_nombre text,                    -- nombre del Excel generado
  salida         text,                    -- stdout crudo del script (opcional)
  creado_en      timestamptz not null default now()
);

-- Índice para listar rápido "mis diseños" del más reciente al más viejo
create index if not exists disenos_usuario_recientes
  on public.disenos (usuario_id, creado_en desc);

-- 2) Activamos RLS en Supabase: sin políticas, nadie toca nada
alter table public.disenos enable row level security;

-- Las 4 políticas siguientes comparan el usuario de la sesión
-- (auth.uid()) con el dueño de la fila (usuario_id):
drop policy if exists "dueno_lee_sus_disenos" on public.disenos;
create policy "dueno_lee_sus_disenos" on public.disenos
  for select using (auth.uid() = usuario_id);        -- leemos solo los nuestros

drop policy if exists "dueno_inserta_sus_disenos" on public.disenos;
create policy "dueno_inserta_sus_disenos" on public.disenos
  for insert with check (auth.uid() = usuario_id);   -- guardamos solo como dueños

drop policy if exists "dueno_actualiza_sus_disenos" on public.disenos;
create policy "dueno_actualiza_sus_disenos" on public.disenos
  for update using (auth.uid() = usuario_id)
  with check (auth.uid() = usuario_id);              -- editamos solo los nuestros

drop policy if exists "dueno_borra_sus_disenos" on public.disenos;
create policy "dueno_borra_sus_disenos" on public.disenos
  for delete using (auth.uid() = usuario_id);        -- borramos solo los nuestros

-- Nota: anon NO tiene ninguna política -> sin sesión no se puede leer ni escribir.

-- ============================================================
-- 4) Supabase Storage: bucket para los informes Excel (Equipo 6)
--    El backend (Render) sube aquí el .xlsx generado y devuelve una
--    URL permanente; así el enlace sobrevive a los reinicios del
--    contenedor. El bucket es público (lectura sin auth); la escritura
--    la hace el backend con la service role key.
-- ============================================================

insert into storage.buckets (id, name, public)
  values ('informes-aci211', 'informes-aci211', true)
  on conflict (id) do nothing;

-- (Opcional) Si el bucket fuera privado, descomentar estas políticas:
--
-- create policy "dueno_lee_sus_informes" on storage.objects
--   for select using (
--     bucket_id = 'informes-aci211'
--     and (storage.foldername(name))[1] = auth.uid()::text
--   );
--
-- create policy "service_role_sube_informes" on storage.objects
--   for insert with check (bucket_id = 'informes-aci211');

-- 3) (Opcional) bucket para archivar los Excel generados por la API
--    Si luego queremos subir el .xlsx a Supabase Storage, descomentamos:
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
