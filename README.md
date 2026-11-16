# Diseño de Mezclas de Concreto — ACI 211

Calculadora de dosificación de concreto (f'c 100–1000 kg/cm²) con el método ACI 211.
El cerebro del cálculo es `diseno_mezclas_aci211.py`: programa de consola interactivo
que genera el informe técnico en Excel (openpyxl). **Ese archivo no se modifica nunca.**

## Arquitectura de producción

| Pieza | Carpeta | Dónde vive |
|---|---|---|
| Frontend React + Vite | `web/frontend` | **Vercel** (Root Directory: `web/frontend`) |
| API FastAPI (ejecuta el cerebro como proceso hijo) | `web/backend` | Contenedor Docker (raíz del repo) en **Railway / Render / Fly** |
| Historial de diseños (Postgres + RLS) | `web/supabase/schema.sql` | **Supabase** (SQL Editor) |

La API expone `POST /api/disenar`, `GET /api/health` y `GET /api/informes/{archivo}`.
En la nube solo atiende `/api/*` (el build del frontend no viaja en la imagen Docker).

## Variables de entorno

| Dónde | Variable | Qué es |
|---|---|---|
| Vercel | `VITE_API_URL` | URL pública de la API, sin barra final |
| Vercel | `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY` | Proyecto Supabase → Settings → API |
| API en Railway/Render | `CORS_ORIGINS` | Dominio de Vercel (separado por comas si hay varios) |

## Desarrollo local

1. `python web/backend/api.py` → API + web servidas en `http://127.0.0.1:8000`
2. `cd web/frontend && npm install && npm run build` (o `npm run dev` para HMR)
3. `web/iniciar_web.bat` hace ambos pasos por ti.

## Regla de oro

La API ejecuta `diseno_mezclas_aci211.py` igual que un usuario tecleando: le envía
las respuestas por stdin y convierte la salida de consola en JSON para el frontend.
