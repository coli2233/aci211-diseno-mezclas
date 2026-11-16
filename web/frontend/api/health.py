# -*- coding: utf-8 -*-
"""Healthcheck de la API serverless en Vercel.

Endpoint:
  GET /api/health -> estado del servicio
"""
from fastapi import FastAPI

app = FastAPI(title="API ACI 211 - health")


@app.get("/api/health")
def salud():
    return {
        "ok": True,
        "programa": "diseno_mezclas_aci211.py",
        "existe": True,
        "modo": "serverless-vercel",
    }
