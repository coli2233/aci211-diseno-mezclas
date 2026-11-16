@echo off
title API ACI 211 + Web React
cd /d "%~dp0backend"
start "" http://127.0.0.1:8000
python api.py
pause
