/* ============================================================
   Cliente de SUPABASE (Equipo 6)
   Este archivo habla directamente con nuestro proyecto Supabase
   "aci211" (https://uiudnidegokdnnisjpho.supabase.co · us-east-1).
   Vercel inyecta VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY en el
   build (las creamos en Vercel → Settings → Environment Variables),
   así que el cliente se configura solo en producción.
   Con él haremos: iniciar sesión (Supabase Auth) y guardar cada
   diseño en la tabla public.disenos que vive en Supabase, protegida
   con RLS (cada usuario solo toca lo suyo).
   Mientras la UI no lo importe, el bundler lo excluye (tree-shaking).
   ============================================================ */

import { createClient } from "@supabase/supabase-js";

/* Leemos las credenciales públicas que Vercel le pasa a Vite */
const URL = import.meta.env.VITE_SUPABASE_URL;
const KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

/* Lo marcamos como "configurado" solo si existen ambas variables */
export const supabaseConfigurado = Boolean(URL && KEY);

/* Creamos el cliente una sola vez y lo compartimos en toda la app;
   si falta configuración dejamos null y la UI avisa en lugar de romper */
export const supabase = supabaseConfigurado
  ? createClient(URL, KEY, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
    })
  : null;

/**
 * Guardamos un diseño en la tabla public.disenos de Supabase.
 * RLS hace el trabajo de seguridad: usuario_id lo rellena auth.uid(),
 * de modo que solo puede guardar quien tenga sesión iniciada y cada
 * fila solo la ve su dueño.
 */
export async function guardarDiseno(datos, resultado) {
  if (!supabase) throw new Error("Supabase no está configurado (revisa VITE_SUPABASE_*).");

  /* Armamos la fila con lo esencial: qué pedimos, qué devolvió el
     script, el nombre del Excel y la salida de consola completa. */
  const fila = {
    descripcion: datos?.descripcion ?? null,
    fc: datos?.fc ?? null,
    datos_entrada: datos,
    dosificacion: resultado?.dosificacion ?? null,
    informe_nombre: resultado?.informe?.nombre ?? null,
    salida: resultado?.salida ?? null,
  };

  /* Enviamos el INSERT a Supabase y pedimos de vuelta id y fecha */
  const { data, error } = await supabase
    .from("disenos")
    .insert(fila)
    .select("id, creado_en")
    .single();

  if (error) throw error;
  return data;
}
