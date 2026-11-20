# -*- coding: utf-8 -*-
"""API REST (FastAPI) para la web React del diseno de mezclas ACI 211
(versión LOCAL / contenedor — Equipo 6).

Lo que hacemos: la API ejecuta diseno_mezclas_aci211.py como proceso
hijo, le envía las respuestas por su entrada estándar (igual que un
usuario tecleando, ENTER = valores por defecto) y convierte la salida
de consola en JSON para el frontend. El cerebro NO lo modificamos.

Dónde corre:
  - En tu PC: python web/backend/api.py (o web/iniciar_web.bat) y sirve
    la web compilada en http://127.0.0.1:8000.
  - En la nube (opcional): la imagen Docker del repo (Railway/Render/Fly)
    atendiendo solo /api/*; el frontal vive en Vercel.
  - En producción ACTUAL: no se usa — la función serverless de
    web/frontend/api/disenar.py hace el mismo trabajo dentro de Vercel.

Endpoints:
  GET  /api/health            -> estado del servicio
  POST /api/disenar           -> ejecuta el script y devuelve los resultados
  GET  /api/informes/{nombre} -> descarga el Excel generado por el script
  GET  /                      -> aplicacion React (build de Vite)
"""
import json
import os
import re
import subprocess
import sys
import threading

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, ConfigDict, ValidationError, field_validator
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# ---------------------------------------------------------------------------
# Validación de entrada (pydantic): rechazamos payloads malos con HTTP 422
# antes de gastar CPU del contenedor ejecutando el cerebro.
# ---------------------------------------------------------------------------
class _Agregado(BaseModel):
    model_config = ConfigDict(extra="ignore")
    puss: float | None = None
    pucs: float | None = None
    pe: float | None = None
    humedad: float | None = None
    absorcion: float | None = None
    mf: float | None = None


class _Probeta(BaseModel):
    model_config = ConfigDict(extra="ignore")
    diametro: float | None = None
    altura: float | None = None
    numero: int | None = None
    desperdicio: float | None = None


class DisenoRequest(BaseModel):
    model_config = ConfigDict(extra="ignore")
    descripcion: str = ""
    fc: float | None = None
    responsable: str = ""
    consistencia: str = "2"
    aire: bool = False
    exposicion: str = "1"
    tipo_cemento: str = ""
    pe_cemento: float | None = None
    arena: _Agregado = _Agregado()
    piedra: _Agregado = _Agregado()
    tmn: str = "3"
    probeta: _Probeta = _Probeta()
    confirmar: bool = True
    informe: bool = True

    @field_validator("fc", "pe_cemento")
    @classmethod
    def _positivo(cls, v):
        if v is not None and v <= 0:
            raise ValueError("debe ser mayor que 0")
        return v

    @field_validator("arena", "piedra", mode="before")
    @classmethod
    def _agregado_dict(cls, v):
        if v is None:
            return _Agregado()
        if isinstance(v, dict):
            return v
        raise ValueError("debe ser un objeto")

    @field_validator("probeta", mode="before")
    @classmethod
    def _probeta_dict(cls, v):
        if v is None:
            return _Probeta()
        if isinstance(v, dict):
            return v
        raise ValueError("debe ser un objeto")


BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_PROYECTO = os.path.dirname(os.path.dirname(BASE))
PROGRAMA = os.path.join(CARPETA_PROYECTO, "diseno_mezclas_aci211.py")
DIST = os.path.join(BASE, "..", "frontend", "dist")

app = FastAPI(title="API ACI 211", version="1.1")

# CORS: en produccion, lista separada por comas en la variable CORS_ORIGINS
# (ej. CORS_ORIGINS=https://aci211.vercel.app). Si no se define, se permite
# cualquier origen porque la web no envia cookies ni credenciales.
_origins = [o.strip() for o in os.environ.get("CORS_ORIGINS", "").split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins or ["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

_candado = threading.Lock()  # serializamos las ejecuciones para que no se pisen

PATRON_INFORME = re.compile(r"^diseno_aci211_\d{8}_\d{6}\.xlsx$")


def _limpiar(valor, defecto=""):
    if valor is None:
        return defecto
    if isinstance(valor, bool):
        return "s" if valor else "n"
    return str(valor).strip()


def _construir_entrada(d):
    """Armamos la lista de respuestas que el script espera por teclado."""
    arena = d.get("arena") or {}
    piedra = d.get("piedra") or {}
    probeta = d.get("probeta") or {}
    lineas = [
        "1",                                                   # opcion del menu
        _limpiar(d.get("descripcion")),                       # descripcion
        _limpiar(d.get("fc")),                                # f'c
        _limpiar(d.get("responsable")),                       # responsable
        _limpiar(d.get("consistencia"), "2"),                 # consistencia 1-3
        "s" if d.get("aire") else "n",                        # aire incorporado
        *([_limpiar(d.get("exposicion"), "1")] if d.get("aire") else []),
        _limpiar(d.get("tipo_cemento")),                      # tipo cemento
        _limpiar(d.get("pe_cemento")),                        # p.e. cemento
        _limpiar(arena.get("puss")),                          # arena PUSS
        _limpiar(arena.get("pucs")),                          # arena PUCS
        _limpiar(arena.get("pe")),                            # arena P.E.
        _limpiar(arena.get("humedad")),                       # arena humedad
        _limpiar(arena.get("absorcion")),                     # arena absorcion
        _limpiar(arena.get("mf")),                            # modulo de fineza
        _limpiar(piedra.get("puss")),                         # piedra PUSS
        _limpiar(piedra.get("pucs")),                         # piedra PUCS
        _limpiar(piedra.get("pe")),                           # piedra P.E.
        _limpiar(piedra.get("humedad")),                      # piedra humedad
        _limpiar(piedra.get("absorcion")),                    # piedra absorcion
        _limpiar(d.get("tmn"), "3"),                          # TMN 1-8
        _limpiar(probeta.get("diametro")),                    # diametro probeta
        _limpiar(probeta.get("altura")),                      # altura probeta
        _limpiar(probeta.get("numero")),                      # numero probetas
        _limpiar(probeta.get("desperdicio")),                 # % desperdicio
    ]
    lineas.append("s" if d.get("confirmar", True) else "n")   # datos correctos?
    if d.get("confirmar", True):
        lineas.append("1" if d.get("informe", True) else "2") # informe?
    lineas.append("")                                          # pausa
    lineas.append("2")                                         # salir del menu
    return "\n".join(lineas) + "\n"


def _buscar_informe(antes, despues):
    nuevos = sorted(set(despues) - set(antes))
    if not nuevos:
        return None
    return os.path.basename(nuevos[-1])


def _pareja(linea):
    trozos = re.split(r"\s{2,}", linea.strip(), maxsplit=1)
    return trozos if len(trozos) == 2 else None


def _seccion(salida, inicio, fin=None):
    """Devuelve el texto entre dos marcas de encabezado."""
    i = salida.find(inicio)
    if i < 0:
        return ""
    i += len(inicio)
    if fin is None:
        return salida[i:]
    j = salida.find(fin, i)
    return salida[i:] if j < 0 else salida[i:j]


def _parsear_datos(salida):
    bloque = _seccion(salida, "RESUMEN DE DATOS INGRESADOS",
                      "Los datos son correctos?")
    datos = {}
    for linea in bloque.splitlines():
        linea = linea.strip()
        if not linea or "=" * 4 in linea or linea.upper().startswith("DATO"):
            continue
        pareja = _pareja(linea)
        if pareja:
            datos[pareja[0].strip()] = pareja[1].strip()
    return datos


def _parsear_dosificacion(salida):
    r = lambda patron: re.search(patron, salida)
    bloque_saco = _seccion(salida, "VOLUMEN POR SACO DE CEMENTO",
                           "CANTIDADES PARA EL VOLUMEN")
    bloque_prob = _seccion(salida, "CANTIDADES PARA EL VOLUMEN CALCULADO")
    slump = r(r"Consistencia\s*:\s*(\S+)\s*\(([^)]+)\)")
    prop = r(r"1\s*:\s*([\d.]+)\s*:\s*([\d.]+)\s+a/c\s*=\s*([\d.]+)")

    def primera(texto, patron, grupo=1, defecto=None):
        m = re.search(patron, texto)
        return m.group(grupo) if m else defecto

    return {
        "fc": primera(salida, r"Resistencia f'c\s*:\s*([\d.]+)"),
        "fcr": primera(salida, r"Resistencia f'cr\s*:\s*([\d.]+)"),
        "ac": primera(salida, r"Relacion a/c\s*:\s*([\d.]+)"),
        "consistencia": slump.group(1) if slump else None,
        "slump_rango": slump.group(2) if slump else None,
        "cemento": primera(salida, r"Cemento \.+?\s*([\d.]+) kg/m3\s*\("),
        "sacos": primera(salida, r"Cemento \.+?\s*[\d.]+ kg/m3\s*\(([\d.]+) sacos\)"),
        "fino": primera(salida, r"Agregado fino \.+?\s*([\d.]+) kg/m3"),
        "grueso": primera(salida, r"Agregado grueso \.+?\s*([\d.]+) kg/m3"),
        "agua_efectiva": primera(salida, r"Agua efectiva \.+?\s*([\d.]+) lt/m3"),
        "prop_fino": prop.group(1) if prop else None,
        "prop_grueso": prop.group(2) if prop else None,
        "vol_fino_saco": primera(bloque_saco, r"Agregado fino \.+?\s*([\d.]+) m3"),
        "vol_grueso_saco": primera(bloque_saco, r"Agregado grueso \.+?\s*([\d.]+) m3"),
        "agua_saco": primera(bloque_saco, r"Agua \.+?\s*([\d.]+) lt"),
        "v_total": primera(bloque_prob, r"V total = ([\d.]+) m3"),
        "cant_cemento": primera(bloque_prob, r"Cemento \.+?\s*([\d.]+) kg"),
        "num_sacos": primera(bloque_prob, r"\(([\d.]+) sacos\)"),
        "cant_fino": primera(bloque_prob, r"Agregado fino \.+?\s*([\d.]+) kg"),
        "cant_grueso": primera(bloque_prob, r"Agregado grueso \.+?\s*([\d.]+) kg"),
        "cant_agua": primera(bloque_prob, r"Agua efectiva \.+?\s*([\d.]+) lt"),
    }


@app.get("/api/health")
def salud():
    # Healthcheck: aquí confirmamos que el cerebro está en su sitio.
    return {"ok": True, "programa": os.path.basename(PROGRAMA),
            "existe": os.path.exists(PROGRAMA)}


@app.post("/api/disenar")
def disenar(d: dict):
    # Validamos la entrada antes de gastar CPU del contenedor.
    try:
        DisenoRequest.model_validate(d or {})
    except ValidationError as e:
        # e.errors() contiene objetos no serializables (ValueError en "ctx");
        # armamos un detalle limpio y serializable.
        detalle = [{"loc": err.get("loc"), "msg": err.get("msg"), "type": err.get("type")} for err in e.errors()]
        raise HTTPException(status_code=422, detail=detalle)

    # Ruta principal: ejecutamos el cerebro con las respuestas del
    # formulario, detectamos el Excel nuevo que dejo y parseamos su salida.
    if not os.path.exists(PROGRAMA):
        raise HTTPException(500, "No se encontro diseno_mezclas_aci211.py")

    antes = set(str(p) for p in
                (os.path.join(CARPETA_PROYECTO, n)
                 for n in os.listdir(CARPETA_PROYECTO)
                 if PATRON_INFORME.match(n)))

    with _candado:
        try:
            proc = subprocess.run(
                [sys.executable, PROGRAMA],
                cwd=CARPETA_PROYECTO,
                input=_construir_entrada(d),
                capture_output=True, text=True,
                encoding="utf-8", errors="replace", timeout=180,
            )
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "El script tardo demasiado en responder.")

    salida = proc.stdout
    if proc.returncode != 0 or "RESUMEN DE LA DOSIFICACION" not in salida:
        return JSONResponse({"ok": False, "detalle": "El script no completo el diseno.",
                             "salida": salida, "stderr": proc.stderr}, status_code=400)

    despues = set(str(p) for p in
                  (os.path.join(CARPETA_PROYECTO, n)
                   for n in os.listdir(CARPETA_PROYECTO)
                   if PATRON_INFORME.match(n)))
    informe = _buscar_informe(antes, despues)

    return {
        "ok": True,
        "datos": _parsear_datos(salida),
        "dosificacion": _parsear_dosificacion(salida),
        "informe": ({"nombre": informe,
                     "url": "/api/informes/" + informe} if informe else None),
        "salida": salida,
    }


@app.get("/api/informes/{nombre}")
def descargar(nombre: str):
    # Descargamos el Excel que el script dejo en la carpeta del proyecto.
    if not PATRON_INFORME.match(nombre):
        raise HTTPException(400, "Nombre de informe invalido.")
    ruta = os.path.join(CARPETA_PROYECTO, nombre)
    if not os.path.exists(ruta):
        raise HTTPException(404, "El informe ya no existe.")
    return FileResponse(ruta, filename=nombre,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# la app React compilada con Vite (frontend/dist) la servimos al final
if os.path.isdir(DIST):
    app.mount("/", StaticFiles(directory=DIST, html=True), name="web")


if __name__ == "__main__":
    import uvicorn
    print("API ACI 211 en http://127.0.0.1:8000  (Ctrl+C para salir)")
    uvicorn.run(app, host="127.0.0.1", port=8000)
