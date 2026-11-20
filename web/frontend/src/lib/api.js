/* ============================================================
   Capa de servicios · única puerta entre React y la API (Equipo 6)
   Aquí decidimos a quién le hablamos:
   - VITE_API_URL definida → llamamos a una API externa por HTTP
     (p. ej. un contenedor desplegado en Railway).
   - VITE_API_URL vacía (así está en Vercel) → llamamos al mismo
     origen: la función serverless /api/disenar que vive dentro del
     propio dominio de Vercel. Cero CORS, cero servidores aparte.
   ============================================================ */

const BASE = import.meta.env.VITE_API_URL || "";

/* Enviamos el formulario al cerebro y devolvemos su respuesta JSON.
   El cálculo real NUNCA ocurre aquí: lo hace el script original
   diseno_mezclas_aci211.py del lado del servidor. */
export async function disenar(datos) {
  let r;
  const headers = { "Content-Type": "application/json" };
  try {
    // Leemos el token de la sesión guardada por Supabase en localStorage
    const key = Object.keys(localStorage).find((k) => k.includes("-auth-token"));
    if (key) {
      const parsed = JSON.parse(localStorage.getItem(key) || "{}");
      const tok = parsed?.access_token || parsed?.currentSession?.access_token;
      if (tok) headers["Authorization"] = `Bearer ${tok}`;
    }
  } catch { /* sin sesión: el backend guarda en carpeta "anon" */ }
  try {
    r = await fetch(BASE + "/api/disenar", {
      method: "POST",
      headers,
      body: JSON.stringify(datos),
    });
  } catch {
    throw new Error("sin conexión con la API");
  }

  /* Si la API rechaza la petición, mostramos su propio "detalle"
     (el mensaje que ella misma escribió) en vez de uno genérico. */
  if (!r.ok) {
    let detalle = "la API respondió " + r.status;
    try {
      const j = await r.json();
      if (j && j.detalle) detalle = j.detalle;
      else if (Array.isArray(j && j.detail)) {
        // Errores de validación (pydantic) -> mensaje legible
        detalle = j.detail.map((e) => (e.loc || []).join(".") + ": " + (e.msg || "")).join("; ");
      } else if (j && j.detail) {
        detalle = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
      }
      if (j && j.stderr && j.stderr.trim()) {
        detalle = detalle + " | " + j.stderr.trim().slice(0, 200);
      }
    } catch { /* cuerpo no JSON */ }
    const e = new Error(detalle);
    e.status = r.status;
    throw e;
  }

  const j = await r.json();

  /* El informe Excel llega de dos formas según el despliegue:
     - serverless (Vercel): lo recibimos incrustado en base64 y lo
       convertimos en data: URL para que el navegador lo descargue.
     - contenedor (local / Railway): llega una ruta relativa y la
       absolutizamos apuntando a la API que la sirvió. */
  if (j && j.informe && j.informe.base64) {
    j.informe.url =
      "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," +
      j.informe.base64;
  } else if (j && j.informe && j.informe.url && j.informe.url.startsWith("/")) {
    j.informe.url = BASE + j.informe.url;
  }
  return j;
}
