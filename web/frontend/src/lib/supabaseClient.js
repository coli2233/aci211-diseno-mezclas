/* ============================================================
   Cliente de SUPABASE (Equipo 6)
   Este archivo habla directamente con nuestro proyecto Supabase
   "aci211" (https://uiudnidegokdnnisjpho.supabase.co · us-east-1).
   Vercel inyecta VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY en el
   build (las creamos en Vercel → Settings → Environment Variables),
   así que el cliente se configura solo en producción.
   Con él hacemos: iniciar sesión (Supabase Auth), guardar cada
   diseño en la tabla public.disenos y listar/borrar el historial
   "Mis diseños"; todo protegido con RLS (cada usuario toca lo suyo).
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

/* Pedimos el usuario de la sesión actual (lo usa el INSERT para
   cumplir la columna usuario_id y la política RLS) */
async function usuarioActual() {
  const { data } = await supabase.auth.getUser();
  if (!data?.user) throw new Error("Inicia sesión para continuar.");
  return data.user;
}

/**
 * Guardamos un diseño en la tabla public.disenos de Supabase.
 * RLS hace el trabajo de seguridad: usuario_id debe ser auth.uid(),
 * de modo que solo puede guardar quien tenga sesión iniciada y cada
 * fila solo la ve su dueño.
 */
export async function guardarDiseno(datos, resultado) {
  if (!supabase) throw new Error("Supabase no está configurado (revisa VITE_SUPABASE_*).");
  const user = await usuarioActual();

  /* Armamos la fila con lo esencial: qué pedimos, qué devolvió el
     script, el nombre del Excel y la salida de consola completa. */
  const fila = {
    usuario_id: user.id,
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

/**
 * Listamos los diseños del usuario (los más recientes primero).
 * No hace falta filtrar por usuario: la política RLS "dueno_lee"
 * ya limita el SELECT a las filas propias.
 */
export async function listarDisenos() {
  if (!supabase) throw new Error("Supabase no está configurado (revisa VITE_SUPABASE_*).");

  const { data, error } = await supabase
    .from("disenos")
    .select("id, descripcion, fc, datos_entrada, dosificacion, informe_nombre, creado_en")
    .order("creado_en", { ascending: false })
    .limit(50);

  if (error) throw error;
  return data ?? [];
}

/* Borramos un diseño propio (RLS "dueno_borra" impide tocar ajenos).
   Pedimos .select() para saber cuántas filas se borraron: si RLS
   bloqueó la operación (sesión caducada o fila ajena), count=0 y
   lanzamos error en vez de "parecer" que se borró. */
export async function eliminarDiseno(id) {
  if (!supabase) throw new Error("Supabase no está configurado (revisa VITE_SUPABASE_*).");

  const { data: { session } } = await supabase.auth.getSession();
  if (!session) throw new Error("Inicia sesión para eliminar diseños.");

  const { data, error } = await supabase
    .from("disenos")
    .delete()
    .eq("id", id)
    .select("id");

  if (error) throw error;
  if (!data || data.length === 0) {
    throw new Error("No se pudo eliminar: la fila no existe o no tienes permiso (sesión caducada).");
  }
  return data;
}
