"""Tests de integracion de la API FastAPI de ACI 211."""
from fastapi.testclient import TestClient

import api

client = TestClient(api.app)


def test_health():
    r = client.get("/api/health")
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] is True
    assert body["existe"] is True


def test_disenar_valido():
    payload = {
        "descripcion": "losa de concreto",
        "fc": 210,
        "responsable": "Equipo 6",
        "consistencia": "2",
        "aire": False,
        "exposicion": "1",
        "tipo_cemento": "Andino Tipo V",
        "pe_cemento": 3.12,
        "arena": {"puss": 1580, "pucs": 1710, "pe": 2.65, "humedad": 3.5, "absorcion": 1.2, "mf": 2.7},
        "piedra": {"puss": 1550, "pucs": 1660, "pe": 2.62, "humedad": 0.4, "absorcion": 1.0},
        "tmn": "3",
        "probeta": {"diametro": 15, "altura": 30, "numero": 3, "desperdicio": 20},
        "confirmar": True,
        "informe": False,
    }
    r = client.post("/api/disenar", json=payload)
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] is True
    assert body["dosificacion"]["fcr"] == "294"
    assert body["dosificacion"]["ac"] == "0.56"


def test_disenar_invalido_devuelve_422():
    r = client.post("/api/disenar", json={"fc": -5, "arena": "x"})
    assert r.status_code == 422
    body = r.json()
    assert "detail" in body