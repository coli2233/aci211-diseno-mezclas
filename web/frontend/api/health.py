# -*- coding: utf-8 -*-
"""Healthcheck — GET /api/health (Equipo 6).

Este archivo lo monta VERCEL como la función serverless /api/health,
en el mismo dominio que la web. Lo llamamos (nosotros y cualquiera
que pruebe el despliegue) para confirmar en un segundo que la
función Python está viva y que la aplicación subió bien.
"""
from fastapi import FastAPI

app = FastAPI(title="API ACI 211 - health")


@app.get("/api/health")
def salud():
    # Respondemos un mínimo de información de estado.
    return {
        "ok": True,
        "programa": "diseno_mezclas_aci211.py",
        "existe": True,
        "modo": "serverless-vercel",
    }
