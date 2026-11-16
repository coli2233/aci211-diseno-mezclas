/* ============================================================
   Capa de servicios · única puerta entre React y la API Python
   VITE_API_URL  = URL pública de la API desplegada (producción)
   "" (vacío)    = mismo origen (desarrollo local: FastAPI sirve todo)
   ============================================================ */

const BASE = import.meta.env.VITE_API_URL || "";

export async function disenar(datos) {
  let r;
  try {
    r = await fetch(BASE + "/api/disenar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datos),
    });
  } catch {
    throw new Error("sin conexión con la API");
  }

  if (!r.ok) {
    let detalle = "la API respondió " + r.status;
    try {
      const j = await r.json();
      if (j && j.detalle) detalle = j.detalle;
    } catch { /* cuerpo no JSON */ }
    const e = new Error(detalle);
    e.status = r.status;
    throw e;
  }

  const j = await r.json();

  // el informe Excel llega de dos formas segun el despliegue:
  //  - serverless (Vercel): base64 incrustado -> se convierte en data: URL
  //  - contenedor (local / Railway): ruta relativa -> se absolutiza
  if (j && j.informe && j.informe.base64) {
    j.informe.url =
      "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," +
      j.informe.base64;
  } else if (j && j.informe && j.informe.url && j.informe.url.startsWith("/")) {
    j.informe.url = BASE + j.informe.url;
  }
  return j;
}
