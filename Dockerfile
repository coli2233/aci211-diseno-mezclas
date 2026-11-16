# Imagen de la API Python (FastAPI) + el cerebro diseno_mezclas_aci211.py
# (Equipo 6). La construimos solo para la ruta de despliegue en
# contenedor (Railway/Render/Fly): la web React vive en Vercel y esta
# imagen atiende únicamente /api/* (dist/ no viaja, StaticFiles se omite).
FROM python:3.12-slim

WORKDIR /proyecto

# Copiamos el cerebro SIN modificarlo: viaja intacto junto a la API
COPY diseno_mezclas_aci211.py ./
COPY web/backend/requirements.txt ./web/backend/
RUN pip install --no-cache-dir -r web/backend/requirements.txt

# Copiamos el resto de web/ (solo backend gracias a .dockerignore)
COPY web/ ./web/

WORKDIR /proyecto/web/backend
EXPOSE 8000
# La plataforma inyecta PORT (Railway/Render/Fly); en local usamos 8000
CMD uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000}
