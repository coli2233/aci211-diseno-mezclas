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
import urllib.request

from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel, ConfigDict, ValidationError, field_validator
from slowapi import Limiter
from slowapi.errors import RateLimitExceeded
from slowapi.util import get_remote_address
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

    @field_validator("puss", "pucs", "pe", mode="before")
    @classmethod
    def _vacio_a_none(cls, v):
        if v == "" or v is None:
            return None
        return v

    @field_validator("puss", "pucs", mode="after")
    @classmethod
    def _min_positivo(cls, v):
        if v is not None and v <= 0:
            raise ValueError("debe ser mayor que 0")
        return v

    @field_validator("mf", mode="before")
    @classmethod
    def _mf_vacio(cls, v):
        if v == "" or v is None:
            return None
        return v

    @field_validator("mf", mode="after")
    @classmethod
    def _mf_min(cls, v):
        if v is not None and v < 1:
            raise ValueError("debe ser mayor que 0 (módulo de fineza)")
        return v

    @field_validator("humedad", "absorcion", mode="after")
    @classmethod
    def _min_cero(cls, v):
        if v is not None and v < 0:
            raise ValueError("no puede ser negativo")
        return v


class _Probeta(BaseModel):
    model_config = ConfigDict(extra="ignore")
    diametro: float | None = None
    altura: float | None = None
    numero: int | None = None
    desperdicio: float | None = None

    @field_validator("diametro", "altura", "numero", "desperdicio", mode="before")
    @classmethod
    def _vacio_a_none(cls, v):
        if v == "" or v is None:
            return None
        return v

    @field_validator("diametro", "altura", mode="after")
    @classmethod
    def _min_dim(cls, v):
        if v is not None and v <= 0:
            raise ValueError("debe ser mayor que 0")
        return v

    @field_validator("numero", mode="after")
    @classmethod
    def _min_num(cls, v):
        if v is not None and v < 1:
            raise ValueError("debe ser al menos 1")
        return v

    @field_validator("desperdicio", mode="after")
    @classmethod
    def _min_desp(cls, v):
        if v is not None and v < 0:
            raise ValueError("no puede ser negativo")
        return v


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

    @field_validator("fc", "pe_cemento", mode="before")
    @classmethod
    def _vacio_a_none_top(cls, v):
        if v == "" or v is None:
            return None
        return v

    @field_validator("fc", "pe_cemento", mode="after")
    @classmethod
    def _positivo(cls, v):
        if v is not None and v <= 0:
            raise ValueError("debe ser mayor que 0")
        return v

    @field_validator("consistencia", mode="before")
    @classmethod
    def _consist(cls, v):
        if v == "" or v is None:
            return "2"
        v = str(v).strip()
        if v not in ("1", "2", "3"):
            raise ValueError("debe ser 1, 2 o 3")
        return v

    @field_validator("exposicion", mode="before")
    @classmethod
    def _expo(cls, v):
        if v == "" or v is None:
            return "1"
        v = str(v).strip()
        if v not in ("1", "2", "3"):
            raise ValueError("debe ser 1, 2 o 3")
        return v

    @field_validator("tmn", mode="before")
    @classmethod
    def _tmn(cls, v):
        if v == "" or v is None:
            return "3"
        v = str(v).strip()
        if v not in ("1", "2", "3", "4", "5", "6", "7", "8"):
            raise ValueError("debe ser del 1 al 8")
        return v

    @field_validator("arena", "piedra", mode="before")
    @classmethod
    def _agregado_dict(cls, v):
        if v is None or v == "":
            return _Agregado()
        if isinstance(v, dict):
            return v
        raise ValueError("debe ser un objeto")

    @field_validator("probeta", mode="before")
    @classmethod
    def _probeta_dict(cls, v):
        if v is None or v == "":
            return _Probeta()
        if isinstance(v, dict):
            return v
        raise ValueError("debe ser un objeto")


BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_PROYECTO = os.path.dirname(os.path.dirname(BASE))
PROGRAMA = os.path.join(CARPETA_PROYECTO, "diseno_mezclas_aci211.py")
DIST = os.path.join(BASE, "..", "frontend", "dist")

app = FastAPI(title="API ACI 211", version="1.1")

# Rate limiting: protege la CPU gratuita de Render de abuso o bucles.
limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter


@app.exception_handler(RateLimitExceeded)
def _rate_limit_handler(request: Request, exc: RateLimitExceeded):
    return JSONResponse(
        {"ok": False, "detalle": "Demasiadas peticiones. Espera un momento e intenta de nuevo."},
        status_code=429,
    )

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
    # Desde la web el usuario ya reviso los datos al pulsar "Calcular", asi
    # que siempre confirmamos: evitamos que un confirmar=False (por ejemplo
    # por un campo vacio en los defaults editados) aborte el diseno.
    lineas.append("s")                                         # datos correctos?
    lineas.append("1" if d.get("informe", True) else "2")       # informe?
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


# ---------------------------------------------------------------------------
# Persistencia del informe en Supabase Storage (opcional).
# Si SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY estan configurados, el Excel
# se sube al bucket "informes-aci211" y se devuelve una URL permanente.
# Si no, se usa el endpoint efimero /api/informes/{nombre} (comportamiento
# anterior, util en desarrollo local).
# ---------------------------------------------------------------------------
def _env(name):
    return (os.environ.get(name) or "").strip()


def _supabase_url():
    return _env("SUPABASE_URL") or _env("VITE_SUPABASE_URL")


def _service_key():
    return _env("SUPABASE_SERVICE_ROLE_KEY")


def _anon_key():
    return _env("VITE_SUPABASE_ANON_KEY") or _env("SUPABASE_ANON_KEY")


def _usuario_id_desde_token(request):
    """Devuelve el user_id de Supabase si hay un Bearer token valido, o None."""
    auth = request.headers.get("authorization") or ""
    if not auth.lower().startswith("bearer "):
        return None
    token = auth.split(" ", 1)[1].strip()
    url = _supabase_url()
    anon = _anon_key()
    if not url or not anon or not token:
        return None
    try:
        req = urllib.request.Request(f"{url}/auth/v1/user", headers={
            "Authorization": f"Bearer {token}",
            "apikey": anon,
        })
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8"))
            return data.get("id")
    except Exception:
        return None


def _subir_informe_supabase(nombre, ruta_archivo, user_id):
    """Sube el Excel a Supabase Storage y devuelve la URL publica.

    Devuelve None si no esta configurado o si falla (fallback al endpoint local).
    """
    url = _supabase_url()
    key = _service_key()
    if not url or not key or not os.path.exists(ruta_archivo):
        return None
    # Guardamos en la raiz del bucket (nombre unico por timestamp) para
    # evitar problemas de parsing de claves con carpetas anidadas.
    bucket = "informes-aci211"
    path = nombre
    try:
        with open(ruta_archivo, "rb") as f:
            data = f.read()
        req = urllib.request.Request(
            f"{url}/storage/v1/object/{bucket}/{path}",
            data=data,
            method="POST",
        )
        req.add_header("Authorization", f"Bearer {key}")
        req.add_header("apikey", key)
        req.add_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        req.add_header("x-upsert", "true")
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read().decode("utf-8", "replace")
            # Usamos la clave devuelta por Supabase para construir la URL
            # (evita desajustes entre la ruta enviada y la almacenada).
            try:
                info = json.loads(raw) if raw else {}
                key_ret = info.get("Key") or info.get("key") or info.get("path") or path
                # Supabase devuelve "bucket/path"; nos quedamos con el path
                if "/" in key_ret and key_ret.split("/", 1)[0] == bucket:
                    key_ret = key_ret.split("/", 1)[1]
                if key_ret:
                    path = key_ret
            except (json.JSONDecodeError, AttributeError):
                pass
        # El bucket es publico: la URL no caduca.
        return f"{url}/storage/v1/object/public/{bucket}/{path}"
    except Exception:
        return None


@app.post("/api/disenar")
@limiter.limit("10/minute")
def disenar(d: dict, request: Request):
    # Validamos la entrada antes de gastar CPU del contenedor.
    try:
        DisenoRequest.model_validate(d or {})
    except ValidationError as e:
        detalle = [{"loc": err.get("loc"), "msg": err.get("msg"), "type": err.get("type")} for err in e.errors()]
        raise HTTPException(status_code=422, detail=detalle)
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
        detalle = "El script no completo el diseno."
        if proc.stderr and proc.stderr.strip():
            detalle = detalle + " | stderr: " + proc.stderr.strip()[:300]
        return JSONResponse({"ok": False, "detalle": detalle,
                             "salida": salida, "stderr": proc.stderr}, status_code=400)

    despues = set(str(p) for p in
                  (os.path.join(CARPETA_PROYECTO, n)
                   for n in os.listdir(CARPETA_PROYECTO)
                   if PATRON_INFORME.match(n)))
    informe = _buscar_informe(antes, despues)

    # Persistimos el informe en Supabase Storage si esta configurado;
    # si no, usamos el endpoint efimero /api/informes/{nombre}.
    informe_obj = None
    if informe:
        ruta_informe = os.path.join(CARPETA_PROYECTO, informe)
        user_id = _usuario_id_desde_token(request)
        url_storage = _subir_informe_supabase(informe, ruta_informe, user_id)
        if url_storage:
            informe_obj = {"nombre": informe, "url": url_storage}
        else:
            informe_obj = {"nombre": informe, "url": "/api/informes/" + informe}

    return {
        "ok": True,
        "datos": _parsear_datos(salida),
        "dosificacion": _parsear_dosificacion(salida),
        "informe": informe_obj,
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
