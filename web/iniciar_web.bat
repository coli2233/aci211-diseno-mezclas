@echo off
REM ============================================================
REM  Arranque local (Equipo 6)
REM ============================================================
REM  Que hace este archivo:
REM    - Levanta nuestra API local (FastAPI) con el cerebro
REM      diseno_mezclas_aci211.py intacto, escuchando en
REM      http://127.0.0.1:8000.
REM    - Sirve la web React compilada (frontend/dist) en el mismo
REM      origen, para que frontend y backend convivan sin CORS.
REM    - Abre el navegador apuntando al puerto 8000.
REM
REM  A que servidor va cada pieza en produccion:
REM    - La web estatica (HTML/JS/CSS) la sirve VERCEL desde
REM      web/frontend (Root Directory en Vercel).
REM    - El calculo Python lo ejecuta RENDER (contenedor Docker de la
REM      raiz del repo): FastAPI recibe POST /api/disenar y corre el
REM      cerebro como proceso hijo.
REM    - La autenticacion y la tabla "disenos" (RLS) viven en
REM      SUPABASE (Auth + Postgres); el frontend habla con Supabase
REM      directo por REST/PostgREST.
REM    - Los informes Excel se persisten en SUPABASE STORAGE
REM      (bucket informes-aci211); Render los sube y devuelve URL
REM      permanente.
REM
REM  En local NO usamos Vercel/Render/Supabase: todo corre en esta PC
REM  para desarrollo y pruebas. Las variables VITE_API_URL y
REM  CORS_ORIGINS quedan vacias/permisivas.
REM
REM  Produccion: https://aci211-diseno-mezclas.vercel.app (Vercel)
REM              https://aci211-diseno-mezclas.onrender.com (Render)
REM ============================================================
title API ACI 211 + Web React
cd /d "%~dp0backend"
start "" http://127.0.0.1:8000
python api.py
pause