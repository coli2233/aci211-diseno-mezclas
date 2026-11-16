# Imagen de la API Python (FastAPI) + el cerebro diseno_mezclas_aci211.py
# La web React se despliega aparte en Vercel; esta imagen solo atiende /api/*
# (el mount de StaticFiles se omite automáticamente porque dist/ no viaja).
FROM python:3.12-slim

WORKDIR /proyecto

# el cerebro NO se modifica: viaja intacto junto a la API
COPY diseno_mezclas_aci211.py ./
COPY web/backend/requirements.txt ./web/backend/
RUN pip install --no-cache-dir -r web/backend/requirements.txt

# el resto de web/ (solo backend gracias a .dockerignore)
COPY web/ ./web/

WORKDIR /proyecto/web/backend
EXPOSE 8000
# PORT la inyecta la plataforma (Railway/Render/Fly); local usa 8000
CMD uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000}
