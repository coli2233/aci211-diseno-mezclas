/* ============================================================
   Cliente centralizado de Supabase (cliente oficial @supabase/supabase-js)
   Solo se activa si existen VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY.
   Mientras la UI no lo importe, el bundler lo excluye (tree-shaking).
   ============================================================ */

import { createClient } from "@supabase/supabase-js";

const URL = import.meta.env.VITE_SUPABASE_URL;
const KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabaseConfigurado = Boolean(URL && KEY);

export const supabase = supabaseConfigurado
  ? createClient(URL, KEY, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
    })
  : null;

/**
 * Guarda un diseño en la tabla public.disenos (RLS: solo el dueño).
 * Requiere sesión iniciada con Supabase Auth (auth.uid()).
 */
export async function guardarDiseno(datos, resultado) {
  if (!supabase) throw new Error("Supabase no está configurado (revisa VITE_SUPABASE_*).");

  const fila = {
    descripcion: datos?.descripcion ?? null,
    fc: datos?.fc ?? null,
    datos_entrada: datos,
    dosificacion: resultado?.dosificacion ?? null,
    informe_nombre: resultado?.informe?.nombre ?? null,
    salida: resultado?.salida ?? null,
  };

  const { data, error } = await supabase
    .from("disenos")
    .insert(fila)
    .select("id, creado_en")
    .single();

  if (error) throw error;
  return data;
}
