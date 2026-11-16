@echo off
REM ============================================================
REM  Arranque local (Equipo 6)
REM  Aqui levantamos la API local (FastAPI) con el cerebro intacto
REM  y abrimos la web en el navegador. En produccion esto no existe:
REM  ahi lo hace Vercel (web estatica + funciones /api serverless).
REM ============================================================
title API ACI 211 + Web React
cd /d "%~dp0backend"
start "" http://127.0.0.1:8000
python api.py
pause
