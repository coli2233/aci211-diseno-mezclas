"""Tests del parser de la salida del cerebro ACI 211."""
import os

import api


FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "salida_ejemplo.txt")


def _salida():
    with open(FIXTURE, encoding="utf-8") as f:
        return f.read()


def test_parsear_datos_devuelve_dict():
    datos = api._parsear_datos(_salida())
    assert isinstance(datos, dict)
    assert len(datos) > 0


def test_parsear_dosificacion_valores_clave():
    d = api._parsear_dosificacion(_salida())
    assert d["fc"] == "210"
    assert d["fcr"] == "294"
    assert d["ac"] == "0.56"
    assert d["cemento"] is not None
    assert d["fino"] is not None
    assert d["grueso"] is not None
    assert d["agua_efectiva"] is not None


def test_parsear_dosificacion_consistencia():
    d = api._parsear_dosificacion(_salida())
    assert d["consistencia"] is not None
    assert d["slump_rango"] is not None


def test_parsear_dosificacion_volumen_total():
    d = api._parsear_dosificacion(_salida())
    assert d["v_total"] is not None
    assert float(d["v_total"]) > 0