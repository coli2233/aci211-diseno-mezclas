# -*- coding: utf-8 -*-
"""API serverless para VERCEL — POST /api/disenar (Equipo 6).

Lo que hacemos en este archivo:
  - VERCEL lo empaqueta como una función Python y lo expone en
    https://aci211-diseno-mezclas.vercel.app/api/disenar (mismo dominio
    que la web: no hay CORS ni servidores externos).
  - El cerebro diseno_mezclas_aci211.py NO lo tocamos: viaja incrustado
    en base64 (SRC_B64, al final) y verificamos su hash al generarlo.
  - En cada petición lo decodificamos a /tmp, lo ejecutamos en memoria
    con un teclado simulado (las mismas respuestas que la API local le
    envía por stdin) y convertimos su salida en JSON.
  - El Excel que el cerebro genera lo leemos y lo devolvemos en base64;
    el navegador lo descarga sin que exista ninguna URL pública.

Diferencia con web/backend/api.py (local/contenedor): no usamos
subprocess ni servimos archivos; todo ocurre dentro de la función.
"""
import base64
import builtins
import contextlib
import io
import os
import re
import shutil
import threading
import uuid

from fastapi import FastAPI
from fastapi.responses import JSONResponse

app = FastAPI(title="API ACI 211 (serverless)", version="1.2")

# Con computo fluido pueden llegar peticiones concurrentes a la misma
# instancia; redirect_stdout es global, así que serializamos la ejecución.
_candado = threading.Lock()

# Patrón del informe: diseno_aci211_YYYYMMDD_HHMMSS.xlsx (lo nombra el cerebro)
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
    # que siempre confirmamos.
    lineas.append("s")                                         # datos correctos?
    lineas.append("1" if d.get("informe", True) else "2")       # informe?
    lineas.append("")                                          # pausa
    lineas.append("2")                                         # salir del menu
    return "\n".join(lineas) + "\n"


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


def _ejecutar_cerebro(d):
    """Ejecutamos el cerebro como si un usuario tecleara.

    Devolvemos (salida, informe, error); informe = None si no hubo Excel.
    """
    # Abrimos un directorio único por petición y decodificamos ahí el
    # cerebro: el script escribe el Excel junto a su propio archivo
    # (__file__), así que cada informe queda aislado de otros usuarios
    # y de restos de instancias tibias.
    run_dir = os.path.join("/tmp", "aci211_" + uuid.uuid4().hex)
    os.makedirs(run_dir, exist_ok=True)
    ruta_py = os.path.join(run_dir, "diseno_mezclas_aci211.py")
    fuente = base64.b64decode(SRC_B64)
    with open(ruta_py, "wb") as f:
        f.write(fuente)

    it = iter(_construir_entrada(d).split("\n"))
    buf = io.StringIO()

    def fake_input(prompt=""):
        # Simulamos input() idéntico a la tubería de la API local:
        # escribimos el mensaje en la salida capturada y, si se agotan
        # las respuestas, lanzamos EOFError (mismo final que en local).
        if prompt:
            buf.write(str(prompt))
        try:
            return next(it)
        except StopIteration:
            raise EOFError("entrada terminada")

    # Preparamos el entorno de ejecución: __name__=__main__,
    # __file__ apuntando al temporal y el input real sustituido por el nuestro
    ns = {
        "__name__": "__main__",
        "__file__": ruta_py,
        "__builtins__": dict(builtins.__dict__, input=fake_input),
    }
    error = None
    with _candado:
        try:
            with contextlib.redirect_stdout(buf):
                exec(compile(fuente.decode("utf-8"), ruta_py, "exec"), ns)
        except SystemExit:
            pass
        except EOFError:
            error = "El programa pidio mas datos de los esperados."
        except Exception as e:  # cualquier fallo se informa como en local
            error = type(e).__name__ + ": " + str(e)
    salida = buf.getvalue()

    # Si el resumen existe, buscamos el Excel más reciente, lo leemos
    # y lo devolvemos incrustado en base64 (Vercel no guarda archivos).
    informe = None
    if error is None and "RESUMEN DE LA DOSIFICACION" in salida:
        candidatos = [n for n in os.listdir(run_dir) if PATRON_INFORME.match(n)]
        if candidatos:
            nombre = sorted(candidatos)[-1]  # el sello temporal ordena solo
            with open(os.path.join(run_dir, nombre), "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            informe = {"nombre": nombre, "url": None, "base64": b64}

    shutil.rmtree(run_dir, ignore_errors=True)
    return salida, informe, error


@app.post("/api/disenar")
def disenar(d: dict):
    # VERCEL expone esta ruta en el dominio de la web. Ejecutamos el
    # cerebro y devolvemos datos + dosificación + informe + salida.
    salida, informe, error = _ejecutar_cerebro(d or {})
    if error is not None or "RESUMEN DE LA DOSIFICACION" not in salida:
        return JSONResponse(
            {
                "ok": False,
                "detalle": error or "El script no completo el diseno.",
                "salida": salida,
            },
            status_code=400,
        )
    return {
        "ok": True,
        "datos": _parsear_datos(salida),
        "dosificacion": _parsear_dosificacion(salida),
        "informe": informe,
        "salida": salida,
    }


# ---------------------------------------------------------------------------
# Fuente original del cerebro: la generamos desde diseno_mezclas_aci211.py
# y la verificamos por hash. NO la editamos a mano: se regenera.
# ---------------------------------------------------------------------------
SRC_B64 = (
    "aW1wb3J0IG1hdGgKaW1wb3J0IG9zCmltcG9ydCByZQppbXBvcnQgc3lzCmZyb20gZGF0ZXRpbWUgaW1wb3J0IGRhdGV0aW1l"
    "CmZyb20gZGVjaW1hbCBpbXBvcnQgUk9VTkRfSEFMRl9VUCwgRGVjaW1hbAoKaW1wb3J0IG51bXB5IGFzIG5wCmltcG9ydCBw"
    "YW5kYXMgYXMgcGQKZnJvbSBvcGVucHl4bCBpbXBvcnQgV29ya2Jvb2sKZnJvbSBvcGVucHl4bC5jaGFydCBpbXBvcnQgQmFy"
    "Q2hhcnQsIFBpZUNoYXJ0LCBSZWZlcmVuY2UKZnJvbSBvcGVucHl4bC5zdHlsZXMgaW1wb3J0IEFsaWdubWVudCwgQm9yZGVy"
    "LCBGb250LCBQYXR0ZXJuRmlsbCwgU2lkZQpmcm9tIG9wZW5weXhsLnV0aWxzIGltcG9ydCBnZXRfY29sdW1uX2xldHRlcgoK"
    "dHJ5OgogICAgc3lzLnN0ZG91dC5yZWNvbmZpZ3VyZShlbmNvZGluZz0idXRmLTgiKQpleGNlcHQgRXhjZXB0aW9uOgogICAg"
    "cGFzcwoKClBFU09fU0FDT19DRU1FTlRPID0gNDIuNQpQRVNPX0VTUEVDSUZJQ09fQUdVQSA9IDEuMDAKUkVTUE9OU0FCTEVf"
    "REVGRUNUTyA9ICJJTkcuIEZMQVZJTyBDRVNBUiBUUlVKSUxMTyBCQUlMT04iCkFOQ0hPID0gNzgKClRNTl9ESVNFTk8gPSBb"
    "JzMvOCInLCAnMS8yIicsICczLzQiJywgJzEiJywgJzEgMS8yIicsICcyIicsICczIicsICc2IiddCk9SREVOX0FJUkVfQVRS"
    "QVBBRE8gPSBbJzMvOCInLCAnMS8yIicsICczLzQiJywgJzEiJywgJzEgMS8yIicsICcyIicsICczIicsICc2IiddCgpUQUJM"
    "QV8wMV9BR1VBID0gewogICAgInNpbl9haXJlIjogewogICAgICAgICcxIiBhIDIiJzogeyczLzgiJzogMjA3LCAnMS8yIic6"
    "IDE5OSwgJzMvNCInOiAxOTAsICcxIic6IDE3OSwKICAgICAgICAgICAgICAgICAgICAnMSAxLzIiJzogMTY2LCAnMiInOiAx"
    "NTQsICczIic6IDEzMCwgJzYiJzogMTEzfSwKICAgICAgICAnMyIgYSA0Iic6IHsnMy84Iic6IDIyOCwgJzEvMiInOiAyMTYs"
    "ICczLzQiJzogMjA1LCAnMSInOiAxOTMsCiAgICAgICAgICAgICAgICAgICAgJzEgMS8yIic6IDE4MSwgJzIiJzogMTY5LCAn"
    "MyInOiAxNDUsICc2Iic6IDEyNH0sCiAgICAgICAgJzYiIGEgNyInOiB7JzMvOCInOiAyNDMsICcxLzIiJzogMjI4LCAnMy80"
    "Iic6IDIxNiwgJzEiJzogMjAyLAogICAgICAgICAgICAgICAgICAgICcxIDEvMiInOiAxOTAsICcyIic6IDE3OCwgJzMiJzog"
    "MTYwLCAnNiInOiBOb25lfSwKICAgIH0sCiAgICAiY29uX2FpcmUiOiB7CiAgICAgICAgJzEiIGEgMiInOiB7JzMvOCInOiAx"
    "ODEsICcxLzIiJzogMTc1LCAnMy80Iic6IDE2OCwgJzEiJzogMTYwLAogICAgICAgICAgICAgICAgICAgICcxIDEvMiInOiAx"
    "NTAsICcyIic6IDE0MiwgJzMiJzogMTIyLCAnNiInOiAxMDd9LAogICAgICAgICczIiBhIDQiJzogeyczLzgiJzogMjAyLCAn"
    "MS8yIic6IDE5MywgJzMvNCInOiAxODQsICcxIic6IDE3NSwKICAgICAgICAgICAgICAgICAgICAnMSAxLzIiJzogMTY1LCAn"
    "MiInOiAxNTcsICczIic6IDEzMywgJzYiJzogMTE5fSwKICAgICAgICAnNiIgYSA3Iic6IHsnMy84Iic6IDIxNiwgJzEvMiIn"
    "OiAyMDUsICczLzQiJzogMTk3LCAnMSInOiAxODQsCiAgICAgICAgICAgICAgICAgICAgJzEgMS8yIic6IDE3NCwgJzIiJzog"
    "MTY2LCAnMyInOiAxNTQsICc2Iic6IE5vbmV9LAogICAgfSwKfQoKVEFCTEFfMDJfQUlSRV9BVFJBUEFETyA9IHsKICAgICcz"
    "LzgiJzogMy4wLCAnMS8yIic6IDIuNSwgJzMvNCInOiAyLjAsICcxIic6IDEuNSwKICAgICcxIDEvMiInOiAxLjAsICcyIic6"
    "IDAuNSwgJzMiJzogMC4zLCAnNCInOiAwLjIsICc2Iic6IDAuMiwKfQoKVEFCTEFfMDNfTUZfQ09NQklOQUNJT04gPSB7CiAg"
    "ICAnMy84Iic6ICAgezY6IDMuOTYsIDc6IDQuMDQsIDg6IDQuMTEsIDk6IDQuMTl9LAogICAgJzEvMiInOiAgIHs2OiA0LjQ2"
    "LCA3OiA0LjU0LCA4OiA0LjYxLCA5OiA0LjY5fSwKICAgICczLzQiJzogICB7NjogNC45NiwgNzogNS4wNCwgODogNS4xMSwg"
    "OTogNS4xOX0sCiAgICAnMSInOiAgICAgezY6IDUuMjYsIDc6IDUuMzQsIDg6IDUuNDEsIDk6IDUuNDl9LAogICAgJzEgMS8y"
    "Iic6IHs2OiA1LjU2LCA3OiA1LjY0LCA4OiA1LjcxLCA5OiA1Ljc5fSwKICAgICcyIic6ICAgICB7NjogNS44NiwgNzogNS45"
    "NCwgODogNi4wMSwgOTogNi4wOX0sCiAgICAnMyInOiAgICAgezY6IDYuMTYsIDc6IDYuMjQsIDg6IDYuMzEsIDk6IDYuMzl9"
    "LAp9CgpUQUJMQV8wNF9BR19HUlVFU08gPSB7CiAgICAnMy84Iic6ICAgezIuNDA6IDAuNTAsIDIuNjA6IDAuNDgsIDIuODA6"
    "IDAuNDYsIDMuMDA6IDAuNDR9LAogICAgJzEvMiInOiAgIHsyLjQwOiAwLjU5LCAyLjYwOiAwLjU3LCAyLjgwOiAwLjU1LCAz"
    "LjAwOiAwLjUzfSwKICAgICczLzQiJzogICB7Mi40MDogMC42NiwgMi42MDogMC42NCwgMi44MDogMC42MiwgMy4wMDogMC42"
    "MH0sCiAgICAnMSInOiAgICAgezIuNDA6IDAuNzEsIDIuNjA6IDAuNjksIDIuODA6IDAuNjcsIDMuMDA6IDAuNjV9LAogICAg"
    "JzEgMS8yIic6IHsyLjQwOiAwLjc2LCAyLjYwOiAwLjc0LCAyLjgwOiAwLjcyLCAzLjAwOiAwLjcwfSwKICAgICcyIic6ICAg"
    "ICB7Mi40MDogMC43OCwgMi42MDogMC43NiwgMi44MDogMC43NCwgMy4wMDogMC43Mn0sCiAgICAnMyInOiAgICAgezIuNDA6"
    "IDAuODEsIDIuNjA6IDAuNzksIDIuODA6IDAuNzcsIDMuMDA6IDAuNzV9LAogICAgJzYiJzogICAgIHsyLjQwOiAwLjg3LCAy"
    "LjYwOiAwLjg1LCAyLjgwOiAwLjgzLCAzLjAwOiAwLjgxfSwKfQoKVEFCTEFfMDVfQUMgPSB7CiAgICAxNTA6ICgwLjgwLCAw"
    "LjcxKSwKICAgIDIwMDogKDAuNzAsIDAuNjEpLAogICAgMjUwOiAoMC42MiwgMC41MyksCiAgICAzMDA6ICgwLjU1LCAwLjQ2"
    "KSwKICAgIDM1MDogKDAuNDgsIDAuNDApLAogICAgNDAwOiAoMC40MywgTm9uZSksCiAgICA0NTA6ICgwLjM4LCBOb25lKSwK"
    "fQoKVEFCTEFfMDZfQUlSRV9JTkNPUlBPUkFETyA9IHsKICAgICczLzgiJzogICAoNC41LCA2LjAsIDcuNSksCiAgICAnMS8y"
    "Iic6ICAgKDQuMCwgNS41LCA3LjApLAogICAgJzMvNCInOiAgICgzLjUsIDUuMCwgNi41KSwKICAgICcxIic6ICAgICAoMy4w"
    "LCA0LjUsIDYuMCksCiAgICAnMSAxLzIiJzogKDIuNSwgNC41LCA1LjUpLAogICAgJzIiJzogICAgICgyLjAsIDQuMCwgNS4w"
    "KSwKICAgICczIic6ICAgICAoMS41LCAzLjUsIDQuNSksCiAgICAnNiInOiAgICAgKDEuMCwgMy4wLCA0LjApLAp9CgpUQUJM"
    "QV8wN19FWFBPU0lDSU9OID0gWwogICAgeyJjb25kaWNpb24iOiAiQmFqYSBwZXJtZWFiaWxpZGFkIC0gZXhwdWVzdG8gYSBh"
    "Z3VhIGR1bGNlIiwKICAgICAid19jX21heCI6IDAuNTAsICJmX2NfbWluX2xpdmlhbm8iOiAyLjYwfSwKICAgIHsiY29uZGlj"
    "aW9uIjogIkJhamEgcGVybWVhYmlsaWRhZCAtIGFndWEgZGUgbWFyIG8gYWd1YXMgc29sdWJsZXMiLAogICAgICJ3X2NfbWF4"
    "IjogMC40NSwgImZfY19taW5fbGl2aWFubyI6IE5vbmV9LAogICAgeyJjb25kaWNpb24iOiAiQmFqYSBwZXJtZWFiaWxpZGFk"
    "IC0gYWd1YXMgY2xvYWNhbGVzIiwKICAgICAid19jX21heCI6IDAuNDUsICJmX2NfbWluX2xpdmlhbm8iOiBOb25lfSwKICAg"
    "IHsiY29uZGljaW9uIjogIkNvbmdlbGFjaW9uIHkgZGVzaGllbG8gLSBiYXJkaW5lbGVzLCBjdW5ldGFzLCBzZWNjLiBkZWxn"
    "YWRhcyIsCiAgICAgIndfY19tYXgiOiAwLjQ1LCAiZl9jX21pbl9saXZpYW5vIjogMzAwfSwKICAgIHsiY29uZGljaW9uIjog"
    "IkNvbmdlbGFjaW9uIHkgZGVzaGllbG8gLSBvdHJvcyBlbGVtZW50b3MiLAogICAgICJ3X2NfbWF4IjogMC41MCwgImZfY19t"
    "aW5fbGl2aWFubyI6IE5vbmV9LAogICAgeyJjb25kaWNpb24iOiAiUHJvdGVjY2lvbiBjb250cmEgY29ycm9zaW9uIChhZ3Vh"
    "IGRlIG1hciAvIHNhbHVicmVzKSIsCiAgICAgIndfY19tYXgiOiAwLjQwLCAiZl9jX21pbl9saXZpYW5vIjogMzI1fSwKICAg"
    "IHsiY29uZGljaW9uIjogIlByb3RlY2Npb24gY29udHJhIGNvcnJvc2lvbiAtIHJlY3VicmltaWVudG8gKzE1IG1tIiwKICAg"
    "ICAid19jX21heCI6IDAuNDUsICJmX2NfbWluX2xpdmlhbm8iOiAzMDB9LApdClRBQkxBXzA3X05PVEEgPSAiTGEgcmVzaXN0"
    "ZW5jaWEgZidjIG5vIGRlYmVyYSBzZXIgbWVub3IgZGUgMjQ1IEtnL2NtMiBwb3IgcmF6b25lcyBkZSBkdXJhYmlsaWRhZCIK"
    "ClRBQkxBXzA4X1BPUkNFTlRBSkVfRklOTyA9IHsKICAgICIyLjMtMi40IjogewogICAgICAgICJyZWRvbmRlYWRvIjogewog"
    "ICAgICAgICAgICAnMy84Iic6IHs1OiA2MCwgNjogNTcsIDc6IDU0LCA4OiA1MX0sICcxLzIiJzogezU6IDQ5LCA2OiA0Niwg"
    "NzogNDMsIDg6IDQwfSwKICAgICAgICAgICAgJzMvNCInOiB7NTogNDEsIDY6IDM4LCA3OiAzNSwgODogMzN9LCAnMSInOiB7"
    "NTogNDAsIDY6IDM3LCA3OiAzNCwgODogMzJ9LAogICAgICAgICAgICAnMSAxLzIiJzogezU6IDM3LCA2OiAzNCwgNzogMzIs"
    "IDg6IDMwfSwgJzIiJzogezU6IDM2LCA2OiAzMywgNzogMzEsIDg6IDI5fSwKICAgICAgICB9LAogICAgICAgICJhbmd1bGFy"
    "IjogewogICAgICAgICAgICAnMy84Iic6IHs1OiA2OSwgNjogNjUsIDc6IDYxLCA4OiA1OH0sICcxLzIiJzogezU6IDU3LCA2"
    "OiA1NCwgNzogNTEsIDg6IDQ4fSwKICAgICAgICAgICAgJzMvNCInOiB7NTogNDgsIDY6IDQ1LCA3OiA0MywgODogNDF9LCAn"
    "MSInOiB7NTogNDcsIDY6IDQ0LCA3OiA0MiwgODogNDB9LAogICAgICAgICAgICAnMSAxLzIiJzogezU6IDQ0LCA2OiA0MSwg"
    "NzogMzksIDg6IDM3fSwgJzIiJzogezU6IDQzLCA2OiA0MCwgNzogMzgsIDg6IDM2fSwKICAgICAgICB9LAogICAgfSwKICAg"
    "ICIyLjYtMi43IjogewogICAgICAgICJyZWRvbmRlYWRvIjogewogICAgICAgICAgICAnMy84Iic6IHs1OiA2NiwgNjogNjIs"
    "IDc6IDU5LCA4OiA1Nn0sICcxLzIiJzogezU6IDUzLCA2OiA1MCwgNzogNDcsIDg6IDQ0fSwKICAgICAgICAgICAgJzMvNCIn"
    "OiB7NTogNDQsIDY6IDQxLCA3OiAzOCwgODogMzZ9LCAnMSInOiB7NTogNDIsIDY6IDM5LCA3OiAzNywgODogMzV9LAogICAg"
    "ICAgICAgICAnMSAxLzIiJzogezU6IDQwLCA2OiAzNywgNzogMzUsIDg6IDMzfSwgJzIiJzogezU6IDM3LCA2OiAzNSwgNzog"
    "MzMsIDg6IDMyfSwKICAgICAgICB9LAogICAgICAgICJhbmd1bGFyIjogewogICAgICAgICAgICAnMy84Iic6IHs1OiA3NSwg"
    "NjogNzEsIDc6IDY3LCA4OiA2NH0sICcxLzIiJzogezU6IDYxLCA2OiA1OCwgNzogNTUsIDg6IDUzfSwKICAgICAgICAgICAg"
    "JzMvNCInOiB7NTogNTEsIDY6IDQ4LCA3OiA0NiwgODogNDR9LCAnMSInOiB7NTogNDksIDY6IDQ2LCA3OiA0NCwgODogNDJ9"
    "LAogICAgICAgICAgICAnMSAxLzIiJzogezU6IDQ3LCA2OiA0NCwgNzogNDIsIDg6IDQwfSwgJzIiJzogezU6IDQ1LCA2OiA0"
    "MiwgNzogNDAsIDg6IDM4fSwKICAgICAgICB9LAogICAgfSwKICAgICIzLjAtMy4xIjogewogICAgICAgICJyZWRvbmRlYWRv"
    "IjogewogICAgICAgICAgICAnMy84Iic6IHs1OiA3NCwgNjogNzAsIDc6IDY2LCA4OiA2Mn0sICcxLzIiJzogezU6IDU5LCA2"
    "OiA1NiwgNzogNTMsIDg6IDUwfSwKICAgICAgICAgICAgJzMvNCInOiB7NTogNDksIDY6IDQ2LCA3OiA0MywgODogNDB9LCAn"
    "MSInOiB7NTogNDcsIDY6IDQ0LCA3OiA0MSwgODogMzh9LAogICAgICAgICAgICAnMSAxLzIiJzogezU6IDQ0LCA2OiA0MSwg"
    "NzogMzgsIDg6IDM2fSwgJzIiJzogezU6IDQyLCA2OiAzOCwgNzogMzYsIDg6IDM0fSwKICAgICAgICB9LAogICAgICAgICJh"
    "bmd1bGFyIjogewogICAgICAgICAgICAnMy84Iic6IHs1OiA4NCwgNjogODAsIDc6IDc2LCA4OiA3M30sICcxLzIiJzogezU6"
    "IDcwLCA2OiA2NiwgNzogNjIsIDg6IDU5fSwKICAgICAgICAgICAgJzMvNCInOiB7NTogNTcsIDY6IDU0LCA3OiA1MSwgODog"
    "NDh9LCAnMSInOiB7NTogNTUsIDY6IDUyLCA3OiA0OSwgODogNDZ9LAogICAgICAgICAgICAnMSAxLzIiJzogezU6IDUyLCA2"
    "OiA0OSwgNzogNDYsIDg6IDQ0fSwgJzIiJzogezU6IDQ5LCA2OiA0NiwgNzogNDQsIDg6IDQyfSwKICAgICAgICB9LAogICAg"
    "fSwKfQoKVEFCTEFfMDlfQUdVQV9QT1JfRk9STUEgPSB7CiAgICAnMy84Iic6ICAgKDE4NSwgMjEyLCAyMDEsIDIyNywgMjMw"
    "LCAyNTApLAogICAgJzEvMiInOiAgICgxODIsIDIwMSwgMTk3LCAyMTYsIDIxOSwgMjM4KSwKICAgICczLzQiJzogICAoMTcw"
    "LCAxODksIDE4NSwgMjA0LCAyMDgsIDIyNyksCiAgICAnMSInOiAgICAgKDE2MywgMTgyLCAxNzgsIDE5NywgMTk3LCAyMTYp"
    "LAogICAgJzEgMS8yIic6ICgxNTUsIDE3MCwgMTcwLCAxODUsIDE4NSwgMjA0KSwKICAgICcyIic6ICAgICAoMTQ4LCAxNjMs"
    "IDE2MywgMTc4LCAxNzgsIDE5NyksCiAgICAnMyInOiAgICAgKDEzNiwgMTUxLCAxNTEsIDE2NywgMTYzLCAxODIpLAp9CgpS"
    "RVNJU1RFTkNJQV9QUk9NRURJT19SRUdMQSA9IFsKICAgICgwLjAsIDIwOS45OTksIDcwLjAsICJGJ2MgKyA3MCIpLAogICAg"
    "KDIxMC4wLCAzNTAuMCwgODQuMCwgIkYnYyArIDg0IiksCiAgICAoMzUwLjAwMSwgOTk5OTkuMCwgOTguMCwgIkYnYyArIDk4"
    "IiksCl0KClRBQkxBX0NPTlNJU1RFTkNJQSA9IFsKICAgICgiU2VjYSIsICAgICAnMCIgYSAyIicsICAgICAnMSIgYSAyIics"
    "ICAiMFwiIGEgMlwiIiksCiAgICAoIlBsw6FzdGljYSIsICczIiBhIDQiJywgICAgICczIiBhIDQiJywgICIzXCIgYSA0XCIi"
    "KSwKICAgICgiRmx1aWRhIiwgICAnPj0gNSInLCAgICAgICAnNiIgYSA3IicsICAiNlwiIGEgN1wiIiksCl0KCkVKRU1QTE9f"
    "Q09SUkVHSURPID0gewogICAgImRlc2NyaXBjaW9uIjogImxvc2EgZGUgY29uY3JldG8iLAogICAgImZjIjogMjEwLAogICAg"
    "InJlc3BvbnNhYmxlIjogUkVTUE9OU0FCTEVfREVGRUNUTywKICAgICJjb25zaXN0ZW5jaWEiOiAiUGzDoXN0aWNhIiwKICAg"
    "ICJjb25fYWlyZSI6IEZhbHNlLAogICAgImV4cG9zaWNpb25faWR4IjogMCwKICAgICJjZW1lbnRvX25vbWJyZSI6ICJBbmRp"
    "bm8gVGlwbyBWIiwKICAgICJwZV9jZW1lbnRvIjogMy4xMiwKICAgICJhcmVuYV9wdXNzIjogMTU4MCwKICAgICJhcmVuYV9w"
    "dWNzIjogMTcxMCwKICAgICJhcmVuYV9wZSI6IDIuNjUsCiAgICAiYXJlbmFfaHVtZWRhZCI6IDMuNSwKICAgICJhcmVuYV9h"
    "YnNvcmNpb24iOiAxLjIsCiAgICAiYXJlbmFfbWYiOiAyLjcsCiAgICAicGllZHJhX3B1c3MiOiAxNTUwLAogICAgInBpZWRy"
    "YV9wdWNzIjogMTY2MCwKICAgICJwaWVkcmFfcGUiOiAyLjYyLAogICAgInBpZWRyYV9odW1lZGFkIjogMC40LAogICAgInBp"
    "ZWRyYV9hYnNvcmNpb24iOiAxLjAsCiAgICAidG1uIjogJzMvNCInLAogICAgInByb2JldGFfZCI6IDE1LAogICAgInByb2Jl"
    "dGFfaCI6IDMwLAogICAgIm5fcHJvYmV0YXMiOiAzLAogICAgImRlc3BlcmRpY2lvIjogMjAsCn0KCgpkZWYgbGltcGlhcl9w"
    "YW50YWxsYSgpOgogICAgb3Muc3lzdGVtKCJjbHMiIGlmIG9zLm5hbWUgPT0gIm50IiBlbHNlICJjbGVhciIpCgpkZWYgbGlu"
    "ZWEoY2FyYWN0ZXI9Ij0iLCBhbmNobz1BTkNITyk6CiAgICBwcmludChjYXJhY3RlciAqIGFuY2hvKQoKZGVmIGVuY2FiZXph"
    "ZG8odGV4dG8pOgogICAgcHJpbnQoKQogICAgbGluZWEoIj0iKQogICAgcHJpbnQodGV4dG8uY2VudGVyKEFOQ0hPKSkKICAg"
    "IGxpbmVhKCI9IikKCmRlZiBwYXVzYShtZW5zYWplPSJcblByZXNpb25lIEVOVEVSIHBhcmEgY29udGludWFyLi4uIik6CiAg"
    "ICBpbnB1dChtZW5zYWplKQoKZGVmIHBlZGlyX3RleHRvKG1lbnNhamUsIGRlZmVjdG89IiIpOgogICAgdmFsb3IgPSBpbnB1"
    "dChtZW5zYWplKS5zdHJpcCgpCiAgICByZXR1cm4gdmFsb3IgaWYgdmFsb3IgZWxzZSBkZWZlY3RvCgpkZWYgcGVkaXJfZmxv"
    "dGFudGUobWVuc2FqZSwgZGVmZWN0bz1Ob25lLCBtaW5pbW89Tm9uZSwgbWF4aW1vPU5vbmUpOgogICAgd2hpbGUgVHJ1ZToK"
    "ICAgICAgICBlbnRyYWRhID0gaW5wdXQobWVuc2FqZSkuc3RyaXAoKS5yZXBsYWNlKCIsIiwgIi4iKQogICAgICAgIGlmIG5v"
    "dCBlbnRyYWRhIGFuZCBkZWZlY3RvIGlzIG5vdCBOb25lOgogICAgICAgICAgICByZXR1cm4gcm91bmQoZmxvYXQoZGVmZWN0"
    "byksIDMpCiAgICAgICAgaWYgbm90IGVudHJhZGE6CiAgICAgICAgICAgIHByaW50KCIgICBbIV0gRGViZSBpbmdyZXNhciB1"
    "biB2YWxvci4iKQogICAgICAgICAgICBjb250aW51ZQogICAgICAgIHRyeToKICAgICAgICAgICAgdmFsb3IgPSBmbG9hdChl"
    "bnRyYWRhKQogICAgICAgIGV4Y2VwdCBWYWx1ZUVycm9yOgogICAgICAgICAgICBwcmludCgiICAgWyFdIFZhbG9yIG5vIHZh"
    "bGlkby4iKQogICAgICAgICAgICBjb250aW51ZQogICAgICAgIGlmIG1pbmltbyBpcyBub3QgTm9uZSBhbmQgdmFsb3IgPCBt"
    "aW5pbW86CiAgICAgICAgICAgIHByaW50KGYiICAgWyFdIERlYmUgc2VyID49IHttaW5pbW99LiIpCiAgICAgICAgICAgIGNv"
    "bnRpbnVlCiAgICAgICAgaWYgbWF4aW1vIGlzIG5vdCBOb25lIGFuZCB2YWxvciA+IG1heGltbzoKICAgICAgICAgICAgcHJp"
    "bnQoZiIgICBbIV0gRGViZSBzZXIgPD0ge21heGltb30uIikKICAgICAgICAgICAgY29udGludWUKICAgICAgICByZXR1cm4g"
    "cm91bmQodmFsb3IsIDMpCgpkZWYgcGVkaXJfb3BjaW9uKG1lbnNhamUsIG9wY2lvbmVzLCBkZWZlY3RvKToKICAgIGZvciBj"
    "bGF2ZSwgKGV0aXF1ZXRhLCBfKSBpbiBvcGNpb25lcy5pdGVtcygpOgogICAgICAgIG1hcmNhID0gIiAgIDwtIHBvciBkZWZl"
    "Y3RvIiBpZiBjbGF2ZSA9PSBzdHIoZGVmZWN0bykgZWxzZSAiIgogICAgICAgIHByaW50KGYiICAgICB7Y2xhdmV9ICB7ZXRp"
    "cXVldGF9e21hcmNhfSIpCiAgICB3aGlsZSBUcnVlOgogICAgICAgIGVudHJhZGEgPSBpbnB1dChtZW5zYWplKS5zdHJpcCgp"
    "CiAgICAgICAgaWYgbm90IGVudHJhZGE6CiAgICAgICAgICAgIHJldHVybiBvcGNpb25lc1tzdHIoZGVmZWN0byldWzFdCiAg"
    "ICAgICAgaWYgZW50cmFkYSBpbiBvcGNpb25lczoKICAgICAgICAgICAgcmV0dXJuIG9wY2lvbmVzW2VudHJhZGFdWzFdCiAg"
    "ICAgICAgcHJpbnQoZiIgICBbIV0gT3BjaW9uIG5vIHZhbGlkYS4gSW5ncmVzZSB1biBudW1lcm8gZGVsIDEgYWwge2xlbihv"
    "cGNpb25lcyl9LiIpCgpkZWYgcGVkaXJfc2lfbm8obWVuc2FqZSwgZGVmZWN0bz0ibiIpOgogICAgd2hpbGUgVHJ1ZToKICAg"
    "ICAgICBlbnRyYWRhID0gaW5wdXQobWVuc2FqZSkuc3RyaXAoKS5sb3dlcigpIG9yIGRlZmVjdG8KICAgICAgICBpZiBlbnRy"
    "YWRhIGluICgicyIsICJzaSIsICJzw60iLCAieSIsICJ5ZXMiKToKICAgICAgICAgICAgcmV0dXJuIFRydWUKICAgICAgICBp"
    "ZiBlbnRyYWRhIGluICgibiIsICJubyIpOgogICAgICAgICAgICByZXR1cm4gRmFsc2UKICAgICAgICBwcmludCgiICAgWyFd"
    "IFJlc3BvbmRhICdzJyBvICduJy4iKQoKZGVmIGludGVycG9sYXIoeCwgeHMsIHlzKToKICAgIHJldHVybiBmbG9hdChucC5p"
    "bnRlcnAoeCwgbnAuYXJyYXkoeHMsIGR0eXBlPWZsb2F0KSwgbnAuYXJyYXkoeXMsIGR0eXBlPWZsb2F0KSkpCgpkZWYgcmFu"
    "Z29fZW50cmVfY2xhdmVzKGNsYXZlcywgeCk6CiAgICBrcyA9IHNvcnRlZChjbGF2ZXMpCiAgICBpZiB4IDw9IGtzWzBdOgog"
    "ICAgICAgIHJldHVybiBrc1swXSwga3NbMF0KICAgIGlmIHggPj0ga3NbLTFdOgogICAgICAgIHJldHVybiBrc1stMV0sIGtz"
    "Wy0xXQogICAgZm9yIGkgaW4gcmFuZ2UobGVuKGtzKSAtIDEpOgogICAgICAgIGlmIGtzW2ldIDw9IHggPD0ga3NbaSArIDFd"
    "OgogICAgICAgICAgICByZXR1cm4ga3NbaV0sIGtzW2kgKyAxXQogICAgcmV0dXJuIGtzWy0xXSwga3NbLTFdCgpkZWYgZm9y"
    "bWF0ZWFyKHZhbG9yLCBkZWNpbWFsZXM9MiwgdmFjaW89Ii0tLSIpOgogICAgaWYgdmFsb3IgaXMgTm9uZToKICAgICAgICBy"
    "ZXR1cm4gdmFjaW8KICAgIHJldHVybiBmInt2YWxvcjoue2RlY2ltYWxlc31mfSIKCmRlZiByZWRvbmRlYXIodmFsb3IsIGRl"
    "Y2ltYWxlcz0yKToKICAgIGlmIHZhbG9yIGlzIE5vbmU6CiAgICAgICAgcmV0dXJuIE5vbmUKICAgIGN1YW50aWEgPSBEZWNp"
    "bWFsKDEpLnNjYWxlYigtZGVjaW1hbGVzKQogICAgYmFzZSA9IERlY2ltYWwoc3RyKHJvdW5kKHZhbG9yLCBkZWNpbWFsZXMg"
    "KyA0KSkpCiAgICByZXR1cm4gZmxvYXQoYmFzZS5xdWFudGl6ZShjdWFudGlhLCByb3VuZGluZz1ST1VORF9IQUxGX1VQKSkK"
    "CgpkZWYgc29saWNpdGFyX2RhdG9zKCk6CiAgICBFSiA9IEVKRU1QTE9fQ09SUkVHSURPCiAgICBlbmNhYmV6YWRvKCJJTkdS"
    "RVNPIERFIERBVE9TIC0gRElTRU5PIERFIE1FWkNMQSBBQ0kgMjExIikKICAgIHByaW50KCJMb3MgdmFsb3JlcyBwb3IgZGVm"
    "ZWN0byBzb24gbG9zIGRlbCBlamVtcGxvIHJlc3VlbHRvICdDT1JSRUdJRE8nLiIpCiAgICBwcmludCgiUHJlc2lvbmUgRU5U"
    "RVIgcGFyYSBhY2VwdGFyIGNhZGEgdmFsb3IgcG9yIGRlZmVjdG8uXG4iKQoKICAgIGQgPSB7fQoKICAgIHByaW50KCItLS0g"
    "REFUT1MgREUgRElTRU5PIC0tLSIpCiAgICBkWyJkZXNjcmlwY2lvbiJdID0gcGVkaXJfdGV4dG8oCiAgICAgICAgZiJEZXNj"
    "cmlwY2lvbiBkZWwgZWxlbWVudG8vcHJveWVjdG8gW3tFSlsnZGVzY3JpcGNpb24nXX1dOiAiLCBFSlsiZGVzY3JpcGNpb24i"
    "XSkKICAgIGRbImZjIl0gPSBwZWRpcl9mbG90YW50ZSgKICAgICAgICBmIlJlc2lzdGVuY2lhIGVzcGVjaWZpY2FkYSBmJ2Mg"
    "W2tnL2NtMl0gW3tFSlsnZmMnXTouMGZ9XTogIiwKICAgICAgICBkZWZlY3RvPUVKWyJmYyJdLCBtaW5pbW89MSkKICAgIGRb"
    "InJlc3BvbnNhYmxlIl0gPSBwZWRpcl90ZXh0bygKICAgICAgICBmIlJlc3BvbnNhYmxlIC8gZG9jZW50ZSBbe0VKWydyZXNw"
    "b25zYWJsZSddfV06ICIsIEVKWyJyZXNwb25zYWJsZSJdKQoKICAgIHByaW50KCJcbkNvbnNpc3RlbmNpYSBkZSBsYSBtZXpj"
    "bGE6IikKICAgIG9wY2lvbmVzX2NvbnNpc3QgPSB7CiAgICAgICAgIjEiOiAoJ1NlY2EgKDAiIGEgMiIpJywgIlNlY2EiKSwK"
    "ICAgICAgICAiMiI6ICgnUGzDoXN0aWNhICgzIiBhIDQiKScsICJQbMOhc3RpY2EiKSwKICAgICAgICAiMyI6ICgnRmx1aWRh"
    "ICg+PSA1IiknLCAiRmx1aWRhIiksCiAgICB9CiAgICBkZWZlY3RvX2NvbnNpc3QgPSBuZXh0KGsgZm9yIGssIHYgaW4gb3Bj"
    "aW9uZXNfY29uc2lzdC5pdGVtcygpIGlmIHZbMV0gPT0gRUpbImNvbnNpc3RlbmNpYSJdKQogICAgZFsiY29uc2lzdGVuY2lh"
    "Il0gPSBwZWRpcl9vcGNpb24oCiAgICAgICAgZiIgICBDb25zaXN0ZW5jaWEgW3tkZWZlY3RvX2NvbnNpc3R9XTogIiwgb3Bj"
    "aW9uZXNfY29uc2lzdCwgZGVmZWN0b19jb25zaXN0KQoKICAgIGRbImNvbl9haXJlIl0gPSBwZWRpcl9zaV9ubygKICAgICAg"
    "ICBmIlxuQ29uY3JldG8gY29uIGFpcmUgaW5jb3Jwb3JhZG8/IChzL24pIFt7J3MnIGlmIEVKWydjb25fYWlyZSddIGVsc2Ug"
    "J24nfV06ICIsCiAgICAgICAgInMiIGlmIEVKWyJjb25fYWlyZSJdIGVsc2UgIm4iKQogICAgaWYgZFsiY29uX2FpcmUiXToK"
    "ICAgICAgICBvcGNpb25lc19leHBvID0geyIxIjogKCJTdWF2ZSIsIDApLCAiMiI6ICgiTW9kZXJhZGEiLCAxKSwgIjMiOiAo"
    "IlNldmVyYSIsIDIpfQogICAgICAgIGRlZmVjdG9fZXhwbyA9IHN0cihFSlsiZXhwb3NpY2lvbl9pZHgiXSArIDEpCiAgICAg"
    "ICAgZXhwb3NpY2lvbiA9IHBlZGlyX29wY2lvbigKICAgICAgICAgICAgZiIgICBUaXBvIGRlIGV4cG9zaWNpb24gW3tkZWZl"
    "Y3RvX2V4cG99XTogIiwgb3BjaW9uZXNfZXhwbywgZGVmZWN0b19leHBvKQogICAgICAgIGRbImV4cG9zaWNpb25faWR4Il0g"
    "PSBleHBvc2ljaW9uCiAgICAgICAgZFsiZXhwb3NpY2lvbl9ub21icmUiXSA9IFsiU3VhdmUiLCAiTW9kZXJhZGEiLCAiU2V2"
    "ZXJhIl1bZXhwb3NpY2lvbl0KICAgIGVsc2U6CiAgICAgICAgZFsiZXhwb3NpY2lvbl9pZHgiXSA9IE5vbmUKICAgICAgICBk"
    "WyJleHBvc2ljaW9uX25vbWJyZSJdID0gIi0iCgogICAgcHJpbnQoIlxuLS0tIENFTUVOVE8gLS0tIikKICAgIGRbImNlbWVu"
    "dG9fbm9tYnJlIl0gPSBwZWRpcl90ZXh0bygKICAgICAgICBmIlRpcG8gZGUgY2VtZW50byBbe0VKWydjZW1lbnRvX25vbWJy"
    "ZSddfV06ICIsIEVKWyJjZW1lbnRvX25vbWJyZSJdKQogICAgZFsicGVfY2VtZW50byJdID0gcGVkaXJfZmxvdGFudGUoCiAg"
    "ICAgICAgZiJQZXNvIGVzcGVjaWZpY28gZGVsIGNlbWVudG8gW3tFSlsncGVfY2VtZW50byddfV06ICIsCiAgICAgICAgZGVm"
    "ZWN0bz1FSlsicGVfY2VtZW50byJdLCBtaW5pbW89MSkKCiAgICBwcmludCgiXG4tLS0gQUdSRUdBRE8gRklOTyAoQVJFTkEp"
    "IC0tLSIpCiAgICBkWyJhcmVuYV9wdXNzIl0gPSBwZWRpcl9mbG90YW50ZSgKICAgICAgICBmIlBlc28gdW5pdGFyaW8gc3Vl"
    "bHRvIHNlY28gW3tFSlsnYXJlbmFfcHVzcyddOi4wZn1dOiAiLAogICAgICAgIGRlZmVjdG89RUpbImFyZW5hX3B1c3MiXSwg"
    "bWluaW1vPTEpCiAgICBkWyJhcmVuYV9wdWNzIl0gPSBwZWRpcl9mbG90YW50ZSgKICAgICAgICBmIlBlc28gdW5pdGFyaW8g"
    "Y29tcGFjdGFkbyBzZWNvIFt7RUpbJ2FyZW5hX3B1Y3MnXTouMGZ9XTogIiwKICAgICAgICBkZWZlY3RvPUVKWyJhcmVuYV9w"
    "dWNzIl0sIG1pbmltbz0xKQogICAgZFsiYXJlbmFfcGUiXSA9IHBlZGlyX2Zsb3RhbnRlKAogICAgICAgIGYiUGVzbyBlc3Bl"
    "Y2lmaWNvIGRlIG1hc2EgW3tFSlsnYXJlbmFfcGUnXX1dOiAiLAogICAgICAgIGRlZmVjdG89RUpbImFyZW5hX3BlIl0sIG1p"
    "bmltbz0xKQogICAgZFsiYXJlbmFfaHVtZWRhZCJdID0gcGVkaXJfZmxvdGFudGUoCiAgICAgICAgZiIlIGRlIGh1bWVkYWQg"
    "W3tFSlsnYXJlbmFfaHVtZWRhZCddfV06ICIsCiAgICAgICAgZGVmZWN0bz1FSlsiYXJlbmFfaHVtZWRhZCJdLCBtaW5pbW89"
    "MCkKICAgIGRbImFyZW5hX2Fic29yY2lvbiJdID0gcGVkaXJfZmxvdGFudGUoCiAgICAgICAgZiIlIGRlIGFic29yY2lvbiBb"
    "e0VKWydhcmVuYV9hYnNvcmNpb24nXX1dOiAiLAogICAgICAgIGRlZmVjdG89RUpbImFyZW5hX2Fic29yY2lvbiJdLCBtaW5p"
    "bW89MCkKICAgIGRbImFyZW5hX21mIl0gPSBwZWRpcl9mbG90YW50ZSgKICAgICAgICBmIk1vZHVsbyBkZSBmaW5lemEgW3tF"
    "SlsnYXJlbmFfbWYnXX1dOiAiLAogICAgICAgIGRlZmVjdG89RUpbImFyZW5hX21mIl0sIG1pbmltbz0xKQoKICAgIHByaW50"
    "KCJcbi0tLSBBR1JFR0FETyBHUlVFU08gKFBJRURSQSkgLS0tIikKICAgIGRbInBpZWRyYV9wdXNzIl0gPSBwZWRpcl9mbG90"
    "YW50ZSgKICAgICAgICBmIlBlc28gdW5pdGFyaW8gc3VlbHRvIHNlY28gW3tFSlsncGllZHJhX3B1c3MnXTouMGZ9XTogIiwK"
    "ICAgICAgICBkZWZlY3RvPUVKWyJwaWVkcmFfcHVzcyJdLCBtaW5pbW89MSkKICAgIGRbInBpZWRyYV9wdWNzIl0gPSBwZWRp"
    "cl9mbG90YW50ZSgKICAgICAgICBmIlBlc28gdW5pdGFyaW8gY29tcGFjdGFkbyBzZWNvIFt7RUpbJ3BpZWRyYV9wdWNzJ106"
    "LjBmfV06ICIsCiAgICAgICAgZGVmZWN0bz1FSlsicGllZHJhX3B1Y3MiXSwgbWluaW1vPTEpCiAgICBkWyJwaWVkcmFfcGUi"
    "XSA9IHBlZGlyX2Zsb3RhbnRlKAogICAgICAgIGYiUGVzbyBlc3BlY2lmaWNvIGRlIG1hc2EgW3tFSlsncGllZHJhX3BlJ119"
    "XTogIiwKICAgICAgICBkZWZlY3RvPUVKWyJwaWVkcmFfcGUiXSwgbWluaW1vPTEpCiAgICBkWyJwaWVkcmFfaHVtZWRhZCJd"
    "ID0gcGVkaXJfZmxvdGFudGUoCiAgICAgICAgZiIlIGRlIGh1bWVkYWQgW3tFSlsncGllZHJhX2h1bWVkYWQnXX1dOiAiLAog"
    "ICAgICAgIGRlZmVjdG89RUpbInBpZWRyYV9odW1lZGFkIl0sIG1pbmltbz0wKQogICAgZFsicGllZHJhX2Fic29yY2lvbiJd"
    "ID0gcGVkaXJfZmxvdGFudGUoCiAgICAgICAgZiIlIGRlIGFic29yY2lvbiBbe0VKWydwaWVkcmFfYWJzb3JjaW9uJ119XTog"
    "IiwKICAgICAgICBkZWZlY3RvPUVKWyJwaWVkcmFfYWJzb3JjaW9uIl0sIG1pbmltbz0wKQoKICAgIHByaW50KCJcblRhbWFu"
    "byBtYXhpbW8gbm9taW5hbCAoVE1OKSBkZWwgYWdyZWdhZG8gZ3J1ZXNvOiIpCiAgICBvcGNpb25lc190bW4gPSB7c3RyKGkg"
    "KyAxKTogKHQsIHQpIGZvciBpLCB0IGluIGVudW1lcmF0ZShUTU5fRElTRU5PKX0KICAgIGRlZmVjdG9fdG1uID0gc3RyKFRN"
    "Tl9ESVNFTk8uaW5kZXgoRUpbInRtbiJdKSArIDEpCiAgICBkWyJ0bW4iXSA9IHBlZGlyX29wY2lvbigKICAgICAgICBmIiAg"
    "IEVsaWphIGVsIFRNTiBbe2RlZmVjdG9fdG1ufSA9IHtFSlsndG1uJ119XTogIiwgb3BjaW9uZXNfdG1uLCBkZWZlY3RvX3Rt"
    "bikKCiAgICBwcmludCgiXG4tLS0gRE9TSUZJQ0FDSU9OIFBBUkEgUFJPQkVUQVMgLS0tIikKICAgIGRbInByb2JldGFfZCJd"
    "ID0gcGVkaXJfZmxvdGFudGUoCiAgICAgICAgZiJEaWFtZXRybyBkZSBwcm9iZXRhIFtjbV0gW3tFSlsncHJvYmV0YV9kJ106"
    "LjBmfV06ICIsCiAgICAgICAgZGVmZWN0bz1FSlsicHJvYmV0YV9kIl0sIG1pbmltbz0xKQogICAgZFsicHJvYmV0YV9oIl0g"
    "PSBwZWRpcl9mbG90YW50ZSgKICAgICAgICBmIkFsdHVyYSBkZSBwcm9iZXRhIFtjbV0gW3tFSlsncHJvYmV0YV9oJ106LjBm"
    "fV06ICIsCiAgICAgICAgZGVmZWN0bz1FSlsicHJvYmV0YV9oIl0sIG1pbmltbz0xKQogICAgZFsibl9wcm9iZXRhcyJdID0g"
    "aW50KHBlZGlyX2Zsb3RhbnRlKAogICAgICAgIGYiTnVtZXJvIGRlIHByb2JldGFzIFt7RUpbJ25fcHJvYmV0YXMnXX1dOiAi"
    "LAogICAgICAgIGRlZmVjdG89RUpbIm5fcHJvYmV0YXMiXSwgbWluaW1vPTEpKQogICAgZFsiZGVzcGVyZGljaW8iXSA9IHBl"
    "ZGlyX2Zsb3RhbnRlKAogICAgICAgIGYiJSBkZSBkZXNwZXJkaWNpbyBbe0VKWydkZXNwZXJkaWNpbyddOi4wZn1dOiAiLAog"
    "ICAgICAgIGRlZmVjdG89RUpbImRlc3BlcmRpY2lvIl0sIG1pbmltbz0wKQoKICAgIHJldHVybiBkCgpkZWYgbW9zdHJhcl9y"
    "ZXN1bWVuX2RhdG9zKGQpOgogICAgZW5jYWJlemFkbygiUkVTVU1FTiBERSBEQVRPUyBJTkdSRVNBRE9TIikKICAgIGZpbGFz"
    "ID0gWwogICAgICAgIFsiUmVzaXN0ZW5jaWEgZidjIChrZy9jbTIpIiwgZiJ7ZFsnZmMnXTouMGZ9Il0sCiAgICAgICAgWyJE"
    "ZXNjcmlwY2lvbiIsIGRbImRlc2NyaXBjaW9uIl1dLAogICAgICAgIFsiUmVzcG9uc2FibGUiLCBkWyJyZXNwb25zYWJsZSJd"
    "XSwKICAgICAgICBbIkNvbnNpc3RlbmNpYSIsIGRbImNvbnNpc3RlbmNpYSJdXSwKICAgICAgICBbIkNvbiBhaXJlIGluY29y"
    "cG9yYWRvIiwgIlNpIiBpZiBkWyJjb25fYWlyZSJdIGVsc2UgIk5vIl0sCiAgICAgICAgWyJFeHBvc2ljaW9uIiwgZFsiZXhw"
    "b3NpY2lvbl9ub21icmUiXV0sCiAgICAgICAgWyJDZW1lbnRvIiwgZFsiY2VtZW50b19ub21icmUiXV0sCiAgICAgICAgWyJQ"
    "LkUuIGNlbWVudG8iLCBmIntkWydwZV9jZW1lbnRvJ106LjJmfSJdLAogICAgICAgIFsiQVJFTkEgLSBQVVNTIChrZy9tMyki"
    "LCBmIntkWydhcmVuYV9wdXNzJ106LjBmfSJdLAogICAgICAgIFsiQVJFTkEgLSBQVUNTIChrZy9tMykiLCBmIntkWydhcmVu"
    "YV9wdWNzJ106LjBmfSJdLAogICAgICAgIFsiQVJFTkEgLSBQLkUuIiwgZiJ7ZFsnYXJlbmFfcGUnXTouMmZ9Il0sCiAgICAg"
    "ICAgWyJBUkVOQSAtIEh1bWVkYWQgKCUpIiwgZiJ7ZFsnYXJlbmFfaHVtZWRhZCddOi4xZn0iXSwKICAgICAgICBbIkFSRU5B"
    "IC0gQWJzb3JjaW9uICglKSIsIGYie2RbJ2FyZW5hX2Fic29yY2lvbiddOi4xZn0iXSwKICAgICAgICBbIkFSRU5BIC0gTW9k"
    "dWxvIGRlIGZpbmV6YSIsIGYie2RbJ2FyZW5hX21mJ106LjFmfSJdLAogICAgICAgIFsiUElFRFJBIC0gUFVTUyAoa2cvbTMp"
    "IiwgZiJ7ZFsncGllZHJhX3B1c3MnXTouMGZ9Il0sCiAgICAgICAgWyJQSUVEUkEgLSBQVUNTIChrZy9tMykiLCBmIntkWydw"
    "aWVkcmFfcHVjcyddOi4wZn0iXSwKICAgICAgICBbIlBJRURSQSAtIFAuRS4iLCBmIntkWydwaWVkcmFfcGUnXTouMmZ9Il0s"
    "CiAgICAgICAgWyJQSUVEUkEgLSBIdW1lZGFkICglKSIsIGYie2RbJ3BpZWRyYV9odW1lZGFkJ106LjFmfSJdLAogICAgICAg"
    "IFsiUElFRFJBIC0gQWJzb3JjaW9uICglKSIsIGYie2RbJ3BpZWRyYV9hYnNvcmNpb24nXTouMWZ9Il0sCiAgICAgICAgWyJQ"
    "SUVEUkEgLSBUTU4iLCBkWyJ0bW4iXV0sCiAgICAgICAgWyJQUk9CRVRBIC0gRGlhbWV0cm8gKGNtKSIsIGYie2RbJ3Byb2Jl"
    "dGFfZCddOi4wZn0iXSwKICAgICAgICBbIlBST0JFVEEgLSBBbHR1cmEgKGNtKSIsIGYie2RbJ3Byb2JldGFfaCddOi4wZn0i"
    "XSwKICAgICAgICBbIlBST0JFVEEgLSBOdW1lcm8iLCBmIntkWyduX3Byb2JldGFzJ119Il0sCiAgICAgICAgWyJQUk9CRVRB"
    "IC0gRGVzcGVyZGljaW8gKCUpIiwgZiJ7ZFsnZGVzcGVyZGljaW8nXTouMGZ9Il0sCiAgICBdCiAgICBwcmludChwZC5EYXRh"
    "RnJhbWUoZmlsYXMsIGNvbHVtbnM9WyJEQVRPIiwgIlZBTE9SIl0pLnRvX3N0cmluZyhpbmRleD1GYWxzZSkpCgoKZGVmIGRp"
    "c2VuYXIoZCk6CiAgICBSID0ge30KICAgIFMgPSBbXQoKICAgIGRlZiBzZWModGl0dWxvLCAqYmxvcXVlcyk6CiAgICAgICAg"
    "aXRlbXMgPSBbXQogICAgICAgIGZvciBiIGluIGJsb3F1ZXM6CiAgICAgICAgICAgIGlmIGIgaXMgTm9uZToKICAgICAgICAg"
    "ICAgICAgIGNvbnRpbnVlCiAgICAgICAgICAgIGlmIGlzaW5zdGFuY2UoYiwgcGQuRGF0YUZyYW1lKToKICAgICAgICAgICAg"
    "ICAgIGl0ZW1zLmFwcGVuZCgoInRhYmxhIiwgYikpCiAgICAgICAgICAgIGVsaWYgaXNpbnN0YW5jZShiLCBzdHIpOgogICAg"
    "ICAgICAgICAgICAgaXRlbXMuYXBwZW5kKCgibm90YSIsIGIpKQogICAgICAgICAgICBlbGlmIGlzaW5zdGFuY2UoYiwgKGxp"
    "c3QsIHR1cGxlKSk6CiAgICAgICAgICAgICAgICBmb3IgeCBpbiBiOgogICAgICAgICAgICAgICAgICAgIGlmIHggaXMgTm9u"
    "ZToKICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICAgICBpZiBpc2luc3RhbmNlKHgs"
    "IHR1cGxlKSBhbmQgbGVuKHgpID09IDIgYW5kIHhbMF0gaW4gKCJ0YWJsYSIsICJub3RhIik6CiAgICAgICAgICAgICAgICAg"
    "ICAgICAgIGl0ZW1zLmFwcGVuZCh4KQogICAgICAgICAgICAgICAgICAgIGVsc2U6CiAgICAgICAgICAgICAgICAgICAgICAg"
    "IGl0ZW1zLmFwcGVuZCgoIm5vdGEiLCBzdHIoeCkpKQogICAgICAgICAgICBlbHNlOgogICAgICAgICAgICAgICAgaXRlbXMu"
    "YXBwZW5kKCgibm90YSIsIHN0cihiKSkpCiAgICAgICAgUy5hcHBlbmQoewogICAgICAgICAgICAidGl0dWxvIjogdGl0dWxv"
    "LAogICAgICAgICAgICAiYmxvcXVlcyI6IGl0ZW1zLAogICAgICAgICAgICAidGFibGEiOiBuZXh0KCh2IGZvciB0LCB2IGlu"
    "IGl0ZW1zIGlmIHQgPT0gInRhYmxhIiksIE5vbmUpLAogICAgICAgICAgICAibm90YXMiOiBbdiBmb3IgdCwgdiBpbiBpdGVt"
    "cyBpZiB0ID09ICJub3RhIl0sCiAgICAgICAgfSkKCiAgICBSWyJ0aXR1bG8iXSA9ICJESVNFw5FPIERFIE1FWkNMQVMgLSBN"
    "w4lUT0RPIENPTUlUw4kgQUNJIDIxMSIKICAgIHByb2JsZW1hID0gKAogICAgICAgIGYiU2UgZGVzZWEgZGlzZcOxYXIgdW5h"
    "IG1lemNsYSBkZSBjb25jcmV0byBwYXJhIHNlciB1dGlsaXphZG8gZW4gbGEgIgogICAgICAgIGYiY29uc3RydWNjacOzbiBk"
    "ZSB1bmEge2RbJ2Rlc2NyaXBjaW9uJ119LiBDb21vIGRhdG8gZGUgZGlzZcOxbyBzZSBzYWJlIHF1ZSAiCiAgICAgICAgZiJs"
    "YSByZXNpc3RlbmNpYSBkZWwgY29uY3JldG8gZXNwZWNpZmljYWRvIHBvciBlbCBwcm95ZWN0aXN0YSBlcyAiCiAgICAgICAg"
    "ZiJmJ2MgPSB7ZFsnZmMnXTouMGZ9IEtnL2NtMi4gTGEgZGlzdHJpYnVjacOzbiBkZWwgw6FyZWEgZGUgYWNlcm8geSBsYXMg"
    "IgogICAgICAgIGYiZGltZW5zaW9uZXMgZGUgbG9zIGVsZW1lbnRvcyBlc3RydWN0dXJhbGVzIGhhY2VuIHJlY29tZW5kYWJs"
    "ZSBlbCBlbXBsZW8gIgogICAgICAgIGYiZGUgdW5hIG1lemNsYSBkZSBjb25zaXN0ZW5jaWEge2RbJ2NvbnNpc3RlbmNpYSdd"
    "Lmxvd2VyKCl9LiBFbCBjZW1lbnRvICIKICAgICAgICBmIlBvcnRsYW5kIGEgZW1wbGVhcnNlIGVzIHtkWydjZW1lbnRvX25v"
    "bWJyZSddfSBjb24gdW4gcGVzbyBlc3BlY8OtZmljbyBkZSAiCiAgICAgICAgZiJ7ZFsncGVfY2VtZW50byddOi4yZn0gZ3Iv"
    "Y20zLiIKICAgICkKICAgIHNlYygiUFJPQkxFTUEiLCBOb25lLCBbcHJvYmxlbWFdKQoKICAgIGNhcmFjdGVyaXN0aWNhcyA9"
    "IHBkLkRhdGFGcmFtZSgKICAgICAgICBbCiAgICAgICAgICAgIFsiUGVzbyBVbml0YXJpbyBTdWVsdG8gU2VjbyIsIGYie2Rb"
    "J2FyZW5hX3B1c3MnXTouMGZ9IEtnL23CsyIsIGYie2RbJ3BpZWRyYV9wdXNzJ106LjBmfSBLZy9twrMiXSwKICAgICAgICAg"
    "ICAgWyJQZXNvIFVuaXRhcmlvIENvbXBhY3RhZG8gU2VjbyIsIGYie2RbJ2FyZW5hX3B1Y3MnXTouMGZ9IEtnL23CsyIsIGYi"
    "e2RbJ3BpZWRyYV9wdWNzJ106LjBmfSBLZy9twrMiXSwKICAgICAgICAgICAgWyJQZXNvIEVzcGVjw61maWNvIGRlIE1hc2Ei"
    "LCBmIntkWydhcmVuYV9wZSddOi4yZn0iLCBmIntkWydwaWVkcmFfcGUnXTouMmZ9Il0sCiAgICAgICAgICAgIFsiQ29udGVu"
    "aWRvIGRlIEh1bWVkYWQiLCBmIntkWydhcmVuYV9odW1lZGFkJ106LjFmfSAlIiwgZiJ7ZFsncGllZHJhX2h1bWVkYWQnXTou"
    "MWZ9ICUiXSwKICAgICAgICAgICAgWyJQb3JjZW50YWplIGRlIEFic29yY2nDs24iLCBmIntkWydhcmVuYV9hYnNvcmNpb24n"
    "XTouMWZ9ICUiLCBmIntkWydwaWVkcmFfYWJzb3JjaW9uJ106LjFmfSAlIl0sCiAgICAgICAgICAgIFsiTcOzZHVsbyBkZSBG"
    "aW5lemEiLCBmIntkWydhcmVuYV9tZiddOi4xZn0iLCAiLS0tLS0tIl0sCiAgICAgICAgICAgIFsiVGFtYcOxbyBNw6F4aW1v"
    "IGRlbCBBZ3JlZ2FkbyIsICItLS0tLS0iLCBkWyJ0bW4iXV0sCiAgICAgICAgXSwKICAgICAgICBjb2x1bW5zPVsiREVTQ1JJ"
    "UENJw5NOIiwgIkFSRU5BIiwgIlBJRURSQSJdLAogICAgKQogICAgc2VjKCJDQVJBQ1RFUsONU1RJQ0FTIERFIExPUyBBR1JF"
    "R0FET1MiLCBjYXJhY3RlcmlzdGljYXMpCgogICAgZmMgPSBkWyJmYyJdCiAgICBmb3IgbG8sIGhpLCBhZGljLCByZWdsYSBp"
    "biBSRVNJU1RFTkNJQV9QUk9NRURJT19SRUdMQToKICAgICAgICBpZiBsbyA8PSBmYyA8PSBoaToKICAgICAgICAgICAgZmNy"
    "ID0gZmMgKyBhZGljCiAgICAgICAgICAgIHJlZ2xhX3VzYWRhID0gcmVnbGEKICAgICAgICAgICAgYnJlYWsKICAgIFJbImZj"
    "Il0gPSBmYwogICAgUlsiZmNyIl0gPSBmY3IKCiAgICB0YWJsYV9mY3IgPSBwZC5EYXRhRnJhbWUoCiAgICAgICAgW1siTWVu"
    "b3MgZGUgMjEwIiwgIkYnYyArIDcwIl0sCiAgICAgICAgIFsiMjEwIGEgMzUwIiwgIkYnYyArIDg0Il0sCiAgICAgICAgIFsi"
    "U29icmUgMzUwIiwgIkYnYyArIDk4Il1dLAogICAgICAgIGNvbHVtbnM9WyJGJ2MgKEtnL2NtwrIpIiwgIkYnY3IiXSwKICAg"
    "ICkKICAgIHNlYygiUEFTTyAxwrA6IERFVEVSTUlOQVIgTEEgUkVTSVNURU5DSUEgUFJPTUVESU8gUEFSQSBFTCBESVNFw5FP"
    "IiwgdGFibGFfZmNyLAogICAgICAgIFtmIkYnY3IgPSB7cmVnbGFfdXNhZGF9ID0ge2ZjOi4wZn0gKyB7YWRpYzouMGZ9IiwK"
    "ICAgICAgICAgZiJGJ2NyID0ge2ZjcjouMGZ9IEtnL2NtwrIiXSkKCiAgICB0YWJsYV9hc2VudCA9IHBkLkRhdGFGcmFtZSgK"
    "ICAgICAgICBbWyJTZWNhIiwgJzAiIGEgMiInXSwgWyJQbMOhc3RpY2EiLCAnMyIgYSA0IiddLCBbIkZsdWlkYSIsICfiiaUg"
    "NSInXV0sCiAgICAgICAgY29sdW1ucz1bIkNPTlNJU1RFTkNJQSIsICJBU0VOVEFNSUVOVE8iXSwKICAgICkKICAgIHNsdW1w"
    "X2NvbCA9IE5vbmUKICAgIGZvciBub21icmUsIHJhbmdvLCBjb2wsIF8gaW4gVEFCTEFfQ09OU0lTVEVOQ0lBOgogICAgICAg"
    "IGlmIG5vbWJyZSA9PSBkWyJjb25zaXN0ZW5jaWEiXToKICAgICAgICAgICAgc2x1bXBfY29sID0gY29sCiAgICAgICAgICAg"
    "IHNsdW1wX3JhbmdvID0gcmFuZ28KICAgICAgICAgICAgYnJlYWsKICAgIFJbInNsdW1wX2NvbCJdID0gc2x1bXBfY29sCiAg"
    "ICBSWyJzbHVtcF9yYW5nbyJdID0gc2x1bXBfcmFuZ28KICAgIHNlYygiUEFTTyAywrA6IERFVEVSTUlOQVIgRUwgQVNFTlRB"
    "TUlFTlRPIiwgdGFibGFfYXNlbnQsCiAgICAgICAgW2YiQ29tbyBsYSBtZXpjbGEgZXMgZGUgY29uc2lzdGVuY2lhIHtkWydj"
    "b25zaXN0ZW5jaWEnXS5sb3dlcigpfSwgIgogICAgICAgICBmImVsIGFzZW50YW1pZW50byBhZGVjdWFkbyBlcyB7c2x1bXBf"
    "cmFuZ299LiJdKQoKICAgIGlkeF9haXJlID0gMSBpZiBkWyJjb25fYWlyZSJdIGVsc2UgMAogICAgZmlsYXNfYWMgPSBbW2Ys"
    "IGZvcm1hdGVhcih2WzBdLCAyKSwgZm9ybWF0ZWFyKHZbMV0sIDIpXQogICAgICAgICAgICAgICAgZm9yIGYsIHYgaW4gc29y"
    "dGVkKFRBQkxBXzA1X0FDLml0ZW1zKCksIHJldmVyc2U9VHJ1ZSldCiAgICB0YWJsYV9hYyA9IHBkLkRhdGFGcmFtZShmaWxh"
    "c19hYywgY29sdW1ucz1bIkYnY3IgKEtnL2NtMikiLCAiU2luIGFpcmUgaW5jb3JwLiIsICJDb24gYWlyZSBpbmNvcnAuIl0p"
    "CgogICAgeHMgPSBbZiBmb3IgZiBpbiBzb3J0ZWQoVEFCTEFfMDVfQUMpIGlmIFRBQkxBXzA1X0FDW2ZdW2lkeF9haXJlXSBp"
    "cyBub3QgTm9uZV0KICAgIHlzID0gW1RBQkxBXzA1X0FDW2ZdW2lkeF9haXJlXSBmb3IgZiBpbiB4c10KICAgIGFjX2V4YWN0"
    "byA9IGludGVycG9sYXIoZmNyLCB4cywgeXMpCiAgICBhYyA9IHJvdW5kKGFjX2V4YWN0bywgMikKICAgIFJbImFjX2V4YWN0"
    "byJdID0gYWNfZXhhY3RvCiAgICBSWyJhYyJdID0gYWMKCiAgICBrMSwgazIgPSByYW5nb19lbnRyZV9jbGF2ZXMoeHMsIGZj"
    "cikKICAgIHYxLCB2MiA9IFRBQkxBXzA1X0FDW2sxXVtpZHhfYWlyZV0sIFRBQkxBXzA1X0FDW2syXVtpZHhfYWlyZV0KICAg"
    "IG5vdGFzX2FjID0gW2YiQ29uIGYnY3IgPSB7ZmNyOi4wZn0gS2cvY23CsiwgaW50ZXJwb2xhbmRvIGVudHJlIHtrMn0gKGEv"
    "Yz17djI6LjJmfSkgIgogICAgICAgICAgICAgICAgZiJ5IHtrMX0gKGEvYz17djE6LjJmfSk6Il0KICAgIGlmIGsxICE9IGsy"
    "OgogICAgICAgIGRlbHRhID0gKGZjciAtIGsxKSAvIChrMiAtIGsxKSAqICh2MiAtIHYxKQogICAgICAgIG5vdGFzX2FjLmFw"
    "cGVuZChmIlgg4oiSIHt2MTouMmZ9ID0gWyh7ZmNyOi4wZn0g4oiSIHtrMX0pIC8gKHtrMn0g4oiSIHtrMX0pXSDDlyAiCiAg"
    "ICAgICAgICAgICAgICAgICAgICAgIGYiKHt2MjouMmZ9IOKIkiB7djE6LjJmfSkgPSAoe2ZjciAtIGsxOi4wZn0ve2syIC0g"
    "azE6LjBmfSkgw5cgIgogICAgICAgICAgICAgICAgICAgICAgICBmIih7djIgLSB2MTouMmZ9KSA9IHtyZWRvbmRlYXIoZGVs"
    "dGEsIDMpOi4zZn0iKQogICAgICAgIG5vdGFzX2FjLmFwcGVuZChmIlggPSB7cmVkb25kZWFyKGRlbHRhLCAzKTouM2Z9ICsg"
    "e3YxOi4yZn0gPSAiCiAgICAgICAgICAgICAgICAgICAgICAgIGYie3JlZG9uZGVhcihhY19leGFjdG8sIDMpOi4zZn0iKQog"
    "ICAgbm90YXNfYWMuYXBwZW5kKGYiUiBhL2MgPSB7YWM6LjJmfSIpCiAgICBzZWMoIlBBU08gM8KwOiBSRUxBQ0nDk04gQUdV"
    "QS1DRU1FTlRPIChhL2MpIiwgdGFibGFfYWMsIG5vdGFzX2FjKQoKICAgIHRpcG8gPSAiY29uX2FpcmUiIGlmIGRbImNvbl9h"
    "aXJlIl0gZWxzZSAic2luX2FpcmUiCiAgICBmaWxhc19hZ3VhX3NpbiA9IFtbcmFuZ29dICsgW2Zvcm1hdGVhcihUQUJMQV8w"
    "MV9BR1VBWyJzaW5fYWlyZSJdW3JhbmdvXVt0XSwgMCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZm9yIHQg"
    "aW4gVE1OX0RJU0VOT10KICAgICAgICAgICAgICAgICAgICAgIGZvciByYW5nbyBpbiBbJzEiIGEgMiInLCAnMyIgYSA0Iics"
    "ICc2IiBhIDciJ11dCiAgICB0YWJsYV9hZ3VhX3NpbiA9IHBkLkRhdGFGcmFtZShmaWxhc19hZ3VhX3NpbiwgY29sdW1ucz1b"
    "IkFzZW50YW1pZW50byJdICsgVE1OX0RJU0VOTykKICAgIGZpbGFzX2FndWFfY29uID0gW1tyYW5nb10gKyBbZm9ybWF0ZWFy"
    "KFRBQkxBXzAxX0FHVUFbImNvbl9haXJlIl1bcmFuZ29dW3RdLCAwKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICBmb3IgdCBpbiBUTU5fRElTRU5PXQogICAgICAgICAgICAgICAgICAgICAgZm9yIHJhbmdvIGluIFsnMSIgYSAyIicsICcz"
    "IiBhIDQiJywgJzYiIGEgNyInXV0KICAgIHRhYmxhX2FndWFfY29uID0gcGQuRGF0YUZyYW1lKGZpbGFzX2FndWFfY29uLCBj"
    "b2x1bW5zPVsiQXNlbnRhbWllbnRvIl0gKyBUTU5fRElTRU5PKQoKICAgIGFndWEgPSBUQUJMQV8wMV9BR1VBW3RpcG9dW3Ns"
    "dW1wX2NvbF1bZFsidG1uIl1dCiAgICBSWyJhZ3VhIl0gPSBmbG9hdChhZ3VhKQogICAgdGlwb190eHQgPSAiY29uIGFpcmUg"
    "aW5jb3Jwb3JhZG8iIGlmIGRbImNvbl9haXJlIl0gZWxzZSAic2luIGFpcmUgaW5jb3Jwb3JhZG8iCiAgICBzZWMoIlBBU08g"
    "NMKwOiBDw4FMQ1VMTyBERSBMQSBDQU5USURBRCBERSBBR1VBIERFIE1FWkNMQURPIFkgQUlSRSIsCiAgICAgICAgIkNvbmNy"
    "ZXRvIHNpbiBhaXJlIGluY29ycG9yYWRvIiwgdGFibGFfYWd1YV9zaW4sCiAgICAgICAgIkNvbmNyZXRvIGNvbiBhaXJlIGlu"
    "Y29ycG9yYWRvIiwgdGFibGFfYWd1YV9jb24sCiAgICAgICAgZiJUTU4gPSB7ZFsndG1uJ119LCBzbHVtcCB7c2x1bXBfcmFu"
    "Z299LCBjb25jcmV0byB7dGlwb190eHR9IOKGkiAiCiAgICAgICAgZiJBZ3VhID0ge2FndWE6LjBmfSBsdC9twrMiKQoKICAg"
    "IGNlbWVudG8gPSBhZ3VhIC8gYWMKICAgIFJbImNlbWVudG9fZGlzZW5vIl0gPSBjZW1lbnRvCiAgICBzZWMoIlBBU08gNcKw"
    "OiBERVRFUk1JTkFSIEVMIENFTUVOVE8iLCBOb25lLAogICAgICAgIFsiUmEvYyA9IGEgLyBjICAg4oaSICAgYyA9IGEgLyBS"
    "YS9jIiwKICAgICAgICAgZiJjID0ge2FndWE6LjBmfSAvIHthYzouMmZ9IiwKICAgICAgICAgZiJjID0ge2NlbWVudG86LjJm"
    "fSBLZy9twrMiXSkKCiAgICBpZiBkWyJjb25fYWlyZSJdOgogICAgICAgIGFpcmVfcGN0ID0gVEFCTEFfMDZfQUlSRV9JTkNP"
    "UlBPUkFET1tkWyJ0bW4iXV1bZFsiZXhwb3NpY2lvbl9pZHgiXV0KICAgICAgICB0YWJsYV9haXJlID0gcGQuRGF0YUZyYW1l"
    "KAogICAgICAgICAgICBbW3QsIGYie1RBQkxBXzA2X0FJUkVfSU5DT1JQT1JBRE9bdF1bMF06LjFmfSAlIiwKICAgICAgICAg"
    "ICAgICBmIntUQUJMQV8wNl9BSVJFX0lOQ09SUE9SQURPW3RdWzFdOi4xZn0gJSIsCiAgICAgICAgICAgICAgZiJ7VEFCTEFf"
    "MDZfQUlSRV9JTkNPUlBPUkFET1t0XVsyXTouMWZ9ICUiXSBmb3IgdCBpbiBUTU5fRElTRU5PXSwKICAgICAgICAgICAgY29s"
    "dW1ucz1bIlRNTiIsICJTdWF2ZSIsICJNb2RlcmFkYSIsICJTZXZlcmEiXSkKICAgICAgICBub3RhID0gKGYiVE1OID0ge2Rb"
    "J3RtbiddfSwgZXhwb3NpY2nDs24ge2RbJ2V4cG9zaWNpb25fbm9tYnJlJ119ICIKICAgICAgICAgICAgICAgIGYi4oaSIEFp"
    "cmUgaW5jb3Jwb3JhZG8gdG90YWwgPSB7YWlyZV9wY3Q6LjBmfSUiKQogICAgICAgIHRpdHVsbzYgPSAiUEFTTyA2wrA6IENP"
    "TlRFTklETyBERSBBSVJFIElOQ09SUE9SQURPIFkgVE9UQUwiCiAgICBlbHNlOgogICAgICAgIGFpcmVfcGN0ID0gVEFCTEFf"
    "MDJfQUlSRV9BVFJBUEFET1tkWyJ0bW4iXV0KICAgICAgICB0YWJsYV9haXJlID0gcGQuRGF0YUZyYW1lKAogICAgICAgICAg"
    "ICBbW3QsIGYie1RBQkxBXzAyX0FJUkVfQVRSQVBBRE9bdF06LjFmfSAlIl0gZm9yIHQgaW4gT1JERU5fQUlSRV9BVFJBUEFE"
    "T10sCiAgICAgICAgICAgIGNvbHVtbnM9WyJUYW1hw7FvIE3DoXhpbW8gTm9taW5hbCIsICJBaXJlIEF0cmFwYWRvIl0pCiAg"
    "ICAgICAgbm90YSA9IGYiVE1OID0ge2RbJ3RtbiddfSDihpIgQWlyZSBBdHJhcGFkbyA9IHthaXJlX3BjdDouMGZ9JSIKICAg"
    "ICAgICB0aXR1bG82ID0gIlBBU08gNsKwOiBDw4FMQ1VMTyBERUwgQUlSRSBBVFJBUEFETyIKICAgIFJbImFpcmVfcGN0Il0g"
    "PSBhaXJlX3BjdAogICAgc2VjKHRpdHVsbzYsIHRhYmxhX2FpcmUsIFtub3RhXSkKCiAgICBmaWxhc19iYm8gPSBbXQogICAg"
    "Zm9yIHQgaW4gVE1OX0RJU0VOTzoKICAgICAgICBmaWxhc19iYm8uYXBwZW5kKFt0XSArIFtmIntUQUJMQV8wNF9BR19HUlVF"
    "U09bdF1bbWZdOi4yZn0iCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZm9yIG1mIGluICgyLjQwLCAyLjYwLCAy"
    "LjgwLCAzLjAwKV0pCiAgICB0YWJsYV9iYm8gPSBwZC5EYXRhRnJhbWUoZmlsYXNfYmJvLCBjb2x1bW5zPVsiVE1OIEFnLiBH"
    "cnVlc28iLCAiTUYgPSAyLjQwIiwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICJN"
    "RiA9IDIuNjAiLCAiTUYgPSAyLjgwIiwgIk1GID0gMy4wMCJdKQoKICAgIG1mID0gZFsiYXJlbmFfbWYiXQogICAgbWZzID0g"
    "c29ydGVkKFRBQkxBXzA0X0FHX0dSVUVTT1tkWyJ0bW4iXV0pCiAgICBiYm8gPSByb3VuZChpbnRlcnBvbGFyKG1mLCBtZnMs"
    "IFtUQUJMQV8wNF9BR19HUlVFU09bZFsidG1uIl1dW21dIGZvciBtIGluIG1mc10pLCAyKQogICAgcGVzb19ncnVlc28gPSBi"
    "Ym8gKiBkWyJwaWVkcmFfcHVjcyJdCiAgICBSWyJiYm8iXSA9IGJibwogICAgUlsicGVzb19ncnVlc28iXSA9IHBlc29fZ3J1"
    "ZXNvCgogICAgazEsIGsyID0gcmFuZ29fZW50cmVfY2xhdmVzKG1mcywgbWYpCiAgICBiMSwgYjIgPSBUQUJMQV8wNF9BR19H"
    "UlVFU09bZFsidG1uIl1dW2sxXSwgVEFCTEFfMDRfQUdfR1JVRVNPW2RbInRtbiJdXVtrMl0KICAgIG5vdGFzX2JibyA9IFtm"
    "IkNvbiBNRiA9IHttZjouMWZ9IHkgVE1OID0ge2RbJ3RtbiddfSwgaW50ZXJwb2xhbmRvIGVudHJlIE1GIHtrMjouMmZ9ICIK"
    "ICAgICAgICAgICAgICAgICBmIih7YjI6LjJmfSkgeSBNRiB7azE6LjJmfSAoe2IxOi4yZn0pOiJdCiAgICBpZiBrMSAhPSBr"
    "MjoKICAgICAgICBkZWx0YV9iYm8gPSAobWYgLSBrMikgLyAoazEgLSBrMikgKiAoYjEgLSBiMikKICAgICAgICBub3Rhc19i"
    "Ym8uYXBwZW5kKGYiWCDiiJIge2IyOi4yZn0gPSBbKHttZjouMWZ9IOKIkiB7azI6LjJmfSkgLyAoe2sxOi4yZn0g4oiSIHtr"
    "MjouMmZ9KV0gw5cgIgogICAgICAgICAgICAgICAgICAgICAgICAgZiIoe2IxOi4yZn0g4oiSIHtiMjouMmZ9KSA9ICh7bWYg"
    "LSBrMjouMWZ9IC8ge2sxIC0gazI6LjFmfSkgw5cgIgogICAgICAgICAgICAgICAgICAgICAgICAgZiJ7YjEgLSBiMjouMmZ9"
    "ID0ge3JlZG9uZGVhcihkZWx0YV9iYm8sIDIpOi4yZn0iKQogICAgICAgIG5vdGFzX2Jiby5hcHBlbmQoZiJYID0ge3JlZG9u"
    "ZGVhcihkZWx0YV9iYm8sIDIpOi4yZn0gKyB7YjI6LjJmfSA9IHtiYm86LjJmfSAgIOKGkiAgICIKICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIGYiYi9ibyA9IHtiYm86LjJmfSIpCiAgICBub3Rhc19iYm8uYXBwZW5kKGYiUGVzbyBBZy4gR3J1ZXNvID0g"
    "e2JibzouMmZ9IMOXIHtkWydwaWVkcmFfcHVjcyddOi4wZn0gPSAiCiAgICAgICAgICAgICAgICAgICAgIGYie3Blc29fZ3J1"
    "ZXNvOi4xZn0gS2cvbcKzIikKICAgIHNlYygiUEFTTyA3wrA6IERFVEVSTUlOQVIgTEEgQ0FOVElEQUQgREUgQUdSRUdBRE8g"
    "R1JVRVNPIiwgdGFibGFfYmJvLCBub3Rhc19iYm8pCgogICAgdm9sX2NlbWVudG8gPSByb3VuZChjZW1lbnRvIC8gKGRbInBl"
    "X2NlbWVudG8iXSAqIDEwMDApLCA0KQogICAgdm9sX2dydWVzbyA9IHJvdW5kKHBlc29fZ3J1ZXNvIC8gKGRbInBpZWRyYV9w"
    "ZSJdICogMTAwMCksIDQpCiAgICB2b2xfYWd1YSA9IHJvdW5kKGFndWEgLyAoUEVTT19FU1BFQ0lGSUNPX0FHVUEgKiAxMDAw"
    "KSwgNCkKICAgIHZvbF9haXJlID0gcm91bmQoYWlyZV9wY3QgLyAxMDAuMCwgNCkKICAgIHZvbF90b3RhbCA9IHJvdW5kKHZv"
    "bF9jZW1lbnRvICsgdm9sX2dydWVzbyArIHZvbF9hZ3VhICsgdm9sX2FpcmUsIDQpCiAgICB2b2xfZmlubyA9IHJvdW5kKDEu"
    "MCAtIHZvbF90b3RhbCwgNCkKICAgIHBlc29fZmlubyA9IHZvbF9maW5vICogZFsiYXJlbmFfcGUiXSAqIDEwMDAKCiAgICBS"
    "LnVwZGF0ZShkaWN0KHZvbF9jZW1lbnRvPXZvbF9jZW1lbnRvLCB2b2xfZ3J1ZXNvPXZvbF9ncnVlc28sIHZvbF9hZ3VhPXZv"
    "bF9hZ3VhLAogICAgICAgICAgICAgICAgICB2b2xfYWlyZT12b2xfYWlyZSwgdm9sX3RvdGFsPXZvbF90b3RhbCwgdm9sX2Zp"
    "bm89dm9sX2Zpbm8sCiAgICAgICAgICAgICAgICAgIHBlc29fZmlubz1wZXNvX2Zpbm8pKQoKICAgIHRhYmxhX3ZvbCA9IHBk"
    "LkRhdGFGcmFtZSgKICAgICAgICBbWyJDZW1lbnRvIiwgZiJ7Y2VtZW50bzouMmZ9IiwgZiJ7ZFsncGVfY2VtZW50byddOi4y"
    "Zn0iLCBmInt2b2xfY2VtZW50bzouM2Z9Il0sCiAgICAgICAgIFsiQWcuIEdydWVzbyIsIGYie3Blc29fZ3J1ZXNvOi4xZn0i"
    "LCBmIntkWydwaWVkcmFfcGUnXTouMmZ9IiwgZiJ7dm9sX2dydWVzbzouM2Z9Il0sCiAgICAgICAgIFsiQWd1YSIsIGYie2Fn"
    "dWE6LjBmfSIsIGYie1BFU09fRVNQRUNJRklDT19BR1VBOi4yZn0iLCBmInt2b2xfYWd1YTouM2Z9Il0sCiAgICAgICAgIFtm"
    "IkFpcmUgKHthaXJlX3BjdDouMGZ9JSkiLCAiLS0tIiwgIi0tLSIsIGYie3ZvbF9haXJlOi4zZn0iXSwKICAgICAgICAgWyJU"
    "T1RBTCIsICItLS0iLCAiLS0tIiwgZiJ7dm9sX3RvdGFsOi4zZn0iXV0sCiAgICAgICAgY29sdW1ucz1bIk1hdGVyaWFsIiwg"
    "IlBlc28gU2VjbyAoS2cpIiwgIlAuRS4iLCAiVm9sLiBBYnNvbHV0byAobcKzKSJdKQogICAgdGFibGFfcGVzb3Nfc2Vjb3Mg"
    "PSBwZC5EYXRhRnJhbWUoCiAgICAgICAgW1siQ2VtZW50byIsIGYie2NlbWVudG86LjJmfSJdLCBbIkFnLiBGaW5vIiwgZiJ7"
    "cGVzb19maW5vOi4zZn0iXSwKICAgICAgICAgWyJBZy4gR3J1ZXNvIiwgZiJ7cGVzb19ncnVlc286LjFmfSJdLCBbIkFndWEi"
    "LCBmInthZ3VhOi4wZn0iXV0sCiAgICAgICAgY29sdW1ucz1bIk1hdGVyaWFsIiwgIlBlc28gU2VjbyAoS2cpIl0pCiAgICBz"
    "ZWMoIlBBU08gOMKwOiBWT0zDmk1FTkVTIEFCU09MVVRPUyDigJQgSEFMTEFSIEVMIEFHUkVHQURPIEZJTk8iLAogICAgICAg"
    "ICJWb2x1bWVuIEFic29sdXRvID0gUGVzbyBTZWNvIC8gKFAuRS4gw5cgMTAwMCkiLAogICAgICAgIHRhYmxhX3ZvbCwKICAg"
    "ICAgICBmIlZvbC4gQWJzb2x1dG8gQWcuIEZpbm8gPSAxIOKIkiB7dm9sX3RvdGFsOi4zZn0gPSB7dm9sX2Zpbm86LjNmfSBt"
    "wrMiLAogICAgICAgIGYiUGVzbyBTZWNvIEFnLiBGaW5vID0ge3ZvbF9maW5vOi4zZn0gw5cge2RbJ2FyZW5hX3BlJ106LjJm"
    "fSDDlyAxMDAwID0ge3Blc29fZmlubzouM2Z9IEtnIiwKICAgICAgICB0YWJsYV9wZXNvc19zZWNvcykKCiAgICBmaW5vX2h1"
    "bWVkbyA9IHBlc29fZmlubyAqIChkWyJhcmVuYV9odW1lZGFkIl0gLyAxMDAuMCArIDEpCiAgICBncnVlc29faHVtZWRvID0g"
    "cGVzb19ncnVlc28gKiAoZFsicGllZHJhX2h1bWVkYWQiXSAvIDEwMC4wICsgMSkKICAgIGZpbm9faHVtZWRvX3JlZCA9IHJv"
    "dW5kKGZpbm9faHVtZWRvKQogICAgZ3J1ZXNvX2h1bWVkb19yZWQgPSByb3VuZChncnVlc29faHVtZWRvKQogICAgUi51cGRh"
    "dGUoZGljdChmaW5vX2h1bWVkbz1maW5vX2h1bWVkbywgZ3J1ZXNvX2h1bWVkbz1ncnVlc29faHVtZWRvLAogICAgICAgICAg"
    "ICAgICAgICBmaW5vX2h1bWVkb19yZWQ9Zmlub19odW1lZG9fcmVkLCBncnVlc29faHVtZWRvX3JlZD1ncnVlc29faHVtZWRv"
    "X3JlZCkpCiAgICBzZWMoIlBBU08gOcKwOiBDT1JSRUNDScOTTiBQT1IgSFVNRURBRCIsIE5vbmUsCiAgICAgICAgWyJQZXNv"
    "IGjDum1lZG8gPSBQZXNvIFNlY28gw5cgKCUgSHVtZWRhZCAvIDEwMCArIDEpIiwKICAgICAgICAgZiJBZy4gRmlubyAgID0g"
    "e3Blc29fZmlubzouM2Z9IMOXICh7ZFsnYXJlbmFfaHVtZWRhZCddOi4xZn0vMTAwICsgMSkgPSB7Zmlub19odW1lZG9fcmVk"
    "fSBLZyIsCiAgICAgICAgIGYiQWcuIEdydWVzbyA9IHtwZXNvX2dydWVzbzouMWZ9IMOXICh7ZFsncGllZHJhX2h1bWVkYWQn"
    "XTouMWZ9LzEwMCArIDEpID0gIgogICAgICAgICBmIntncnVlc29faHVtZWRvOi4yZn0g4omIIHtncnVlc29faHVtZWRvX3Jl"
    "ZH0gS2ciXSkKCiAgICBhZ3VhX2xpYnJlX2Zpbm8gPSBwZXNvX2Zpbm8gKiAoZFsiYXJlbmFfYWJzb3JjaW9uIl0gLSBkWyJh"
    "cmVuYV9odW1lZGFkIl0pIC8gMTAwLjAKICAgIGFndWFfbGlicmVfZ3J1ZXNvID0gcGVzb19ncnVlc28gKiAoZFsicGllZHJh"
    "X2Fic29yY2lvbiJdIC0gZFsicGllZHJhX2h1bWVkYWQiXSkgLyAxMDAuMAogICAgYWd1YV9saWJyZV90b3RhbCA9IHJvdW5k"
    "KGFndWFfbGlicmVfZmlubywgMikgKyByb3VuZChhZ3VhX2xpYnJlX2dydWVzbywgMikKICAgIFIudXBkYXRlKGRpY3QoYWd1"
    "YV9saWJyZV9maW5vPWFndWFfbGlicmVfZmlubywgYWd1YV9saWJyZV9ncnVlc289YWd1YV9saWJyZV9ncnVlc28sCiAgICAg"
    "ICAgICAgICAgICAgIGFndWFfbGlicmVfdG90YWw9YWd1YV9saWJyZV90b3RhbCkpCiAgICBzZWMoIlBBU08gMTDCsDogQ09S"
    "UkVDQ0nDk04gUE9SIEFCU09SQ0nDk04iLCBOb25lLAogICAgICAgIFsiQWd1YSBsaWJyZSA9IFBlc28gU2VjbyDDlyAoJSBB"
    "YnNvcmNpw7NuIOKIkiAlIEh1bWVkYWQpIC8gMTAwIiwKICAgICAgICAgZiJBZy4gRmlubyAgID0ge3Blc29fZmlubzouM2Z9"
    "IMOXICh7ZFsnYXJlbmFfYWJzb3JjaW9uJ106LjFmfSDiiJIge2RbJ2FyZW5hX2h1bWVkYWQnXTouMWZ9KSAvIDEwMCA9ICIK"
    "ICAgICAgICAgZiJ7YWd1YV9saWJyZV9maW5vOi4yZn0gTHQiLAogICAgICAgICBmIkFnLiBHcnVlc28gPSB7cGVzb19ncnVl"
    "c286LjFmfSDDlyAoe2RbJ3BpZWRyYV9hYnNvcmNpb24nXTouMWZ9IOKIkiB7ZFsncGllZHJhX2h1bWVkYWQnXTouMWZ9KSAv"
    "IDEwMCA9ICIKICAgICAgICAgZiIre2FndWFfbGlicmVfZ3J1ZXNvOi4yZn0gTHQiLAogICAgICAgICBmIs6jIEFndWEgTGli"
    "cmUgPSB7cm91bmQoYWd1YV9saWJyZV9maW5vLDIpOi4yZn0gKyB7cm91bmQoYWd1YV9saWJyZV9ncnVlc28sMik6LjJmfSA9"
    "ICIKICAgICAgICAgZiJ7YWd1YV9saWJyZV90b3RhbDouMmZ9IEx0Il0pCgogICAgYWd1YV9lZmVjdGl2YSA9IGFndWEgKyBh"
    "Z3VhX2xpYnJlX3RvdGFsCiAgICBSWyJhZ3VhX2VmZWN0aXZhIl0gPSBhZ3VhX2VmZWN0aXZhCiAgICBzZWMoIlBBU08gMTHC"
    "sDogQUdVQSBFRkVDVElWQSIsIE5vbmUsCiAgICAgICAgWyJBZ3VhIEVmZWN0aXZhID0gQWd1YSBEaXNlw7FvICsgQWd1YSBM"
    "aWJyZSIsCiAgICAgICAgIGYiQWd1YSBFZmVjdGl2YSA9IHthZ3VhOi4wZn0gKyAoe2FndWFfbGlicmVfdG90YWw6LjJmfSkg"
    "PSB7YWd1YV9lZmVjdGl2YTouMmZ9IEx0L23CsyJdKQoKICAgIGNlbWVudG9faHVtZWRvID0gcmVkb25kZWFyKGFndWFfZWZl"
    "Y3RpdmEgLyBhYywgMikKICAgIFJbImNlbWVudG9faHVtZWRvIl0gPSBjZW1lbnRvX2h1bWVkbwogICAgc2VjKCJQQVNPIDEy"
    "wrA6IERJU0XDkU8gVEXDk1JJQ08gSMOaTUVETyIsIE5vbmUsCiAgICAgICAgWyJSYS9jID0gQWd1YSBFZmVjdGl2YSAvIENl"
    "bWVudG8gICDihpIgICBDID0gQWd1YSBFZmVjdGl2YSAvIFJhL2MiLAogICAgICAgICBmIkMgPSB7YWd1YV9lZmVjdGl2YTou"
    "MmZ9IC8ge2FjOi4yZn0gPSB7YWd1YV9lZmVjdGl2YSAvIGFjOi4zZn0iLAogICAgICAgICBmIkMgPSB7Y2VtZW50b19odW1l"
    "ZG86LjJmfSBLZy9twrMiXSkKCiAgICBwcm9wX2Zpbm8gPSBmaW5vX2h1bWVkb19yZWQgLyBjZW1lbnRvX2h1bWVkbwogICAg"
    "cHJvcF9ncnVlc28gPSBncnVlc29faHVtZWRvX3JlZCAvIGNlbWVudG9faHVtZWRvCiAgICBwcm9wX2FndWEgPSBhZ3VhX2Vm"
    "ZWN0aXZhIC8gY2VtZW50b19odW1lZG8KICAgIFIudXBkYXRlKGRpY3QocHJvcF9maW5vPXByb3BfZmlubywgcHJvcF9ncnVl"
    "c289cHJvcF9ncnVlc28sIHByb3BfYWd1YT1wcm9wX2FndWEpKQoKICAgIHRhYmxhX3Byb3AgPSBwZC5EYXRhRnJhbWUoCiAg"
    "ICAgICAgW1siQ2VtZW50byIsIGYie2NlbWVudG9faHVtZWRvOi4yZn0iLCBmIsO3IHtjZW1lbnRvX2h1bWVkbzouMmZ9Iiwg"
    "IjEiXSwKICAgICAgICAgWyJBZy4gRmlubyIsIGYie2Zpbm9faHVtZWRvX3JlZH0iLCBmIsO3IHtjZW1lbnRvX2h1bWVkbzou"
    "MmZ9IiwgZiJ7cHJvcF9maW5vOi4yZn0iXSwKICAgICAgICAgWyJBZy4gR3J1ZXNvIiwgZiJ7Z3J1ZXNvX2h1bWVkb19yZWR9"
    "IiwgZiLDtyB7Y2VtZW50b19odW1lZG86LjJmfSIsIGYie3Byb3BfZ3J1ZXNvOi4yZn0iXSwKICAgICAgICAgWyJBZ3VhIiwg"
    "ZiJ7YWd1YV9lZmVjdGl2YTouMmZ9IiwgZiLDtyB7Y2VtZW50b19odW1lZG86LjJmfSIsIGYiUmEvYyA9IHtwcm9wX2FndWE6"
    "LjJmfSJdXSwKICAgICAgICBjb2x1bW5zPVsiTWF0ZXJpYWwiLCAiUGVzbyBIw7ptZWRvIChLZykiLCAiw7cgQ2VtZW50byIs"
    "ICJQcm9wb3JjacOzbiJdKQogICAgc2VjKCJQQVNPIDEzwrA6IFBST1BPUkNJT05FUyBFTiBQRVNPIiwKICAgICAgICBmIkRp"
    "dmlkaWVuZG8gY2FkYSBtYXRlcmlhbCBlbnRyZSBlbCBwZXNvIGRlbCBjZW1lbnRvICh7Y2VtZW50b19odW1lZG86LjJmfSBL"
    "Zyk6IiwKICAgICAgICB0YWJsYV9wcm9wLAogICAgICAgIGYiUHJvcG9yY2nDs24gZW4gcGVzbzogMSA6IHtwcm9wX2Zpbm86"
    "LjJmfSA6IHtwcm9wX2dydWVzbzouMmZ9IGNvbiBSYS9jID0ge3Byb3BfYWd1YTouMmZ9IikKCiAgICBwcm9wX2Zpbm9fMiA9"
    "IHJlZG9uZGVhcihwcm9wX2Zpbm8sIDIpCiAgICBwcm9wX2dydWVzb18yID0gcmVkb25kZWFyKHByb3BfZ3J1ZXNvLCAyKQog"
    "ICAgcHJvcF9hZ3VhXzIgPSByZWRvbmRlYXIocHJvcF9hZ3VhLCAyKQogICAgdm9sX2Zpbm9fc2FjbyA9IChwcm9wX2Zpbm9f"
    "MiAqIFBFU09fU0FDT19DRU1FTlRPKSAvIGRbImFyZW5hX3B1c3MiXQogICAgdm9sX2dydWVzb19zYWNvID0gKHByb3BfZ3J1"
    "ZXNvXzIgKiBQRVNPX1NBQ09fQ0VNRU5UTykgLyBkWyJwaWVkcmFfcHVzcyJdCiAgICBhZ3VhX3NhY28gPSBwcm9wX2FndWFf"
    "MiAqIFBFU09fU0FDT19DRU1FTlRPCiAgICBSLnVwZGF0ZShkaWN0KHZvbF9maW5vX3NhY289dm9sX2Zpbm9fc2Fjbywgdm9s"
    "X2dydWVzb19zYWNvPXZvbF9ncnVlc29fc2FjbywKICAgICAgICAgICAgICAgICAgYWd1YV9zYWNvPWFndWFfc2FjbykpCgog"
    "ICAgdGFibGFfdm9sX3NhY28gPSBwZC5EYXRhRnJhbWUoCiAgICAgICAgW1siQ2VtZW50byIsICJSZWZlcmVuY2lhIiwgIjEi"
    "LCBmInNhY28gKHtQRVNPX1NBQ09fQ0VNRU5UT30gS2cpIl0sCiAgICAgICAgIFsiQWcuIEZpbm8iLCBmIih7cHJvcF9maW5v"
    "XzI6LjJmfcOXe1BFU09fU0FDT19DRU1FTlRPfSkve2RbJ2FyZW5hX3B1c3MnXTouMGZ9IiwKICAgICAgICAgIGYie3ZvbF9m"
    "aW5vX3NhY286LjNmfSIsICJtwrMiXSwKICAgICAgICAgWyJBZy4gR3J1ZXNvIiwgZiIoe3Byb3BfZ3J1ZXNvXzI6LjJmfcOX"
    "e1BFU09fU0FDT19DRU1FTlRPfSkve2RbJ3BpZWRyYV9wdXNzJ106LjBmfSIsCiAgICAgICAgICBmInt2b2xfZ3J1ZXNvX3Nh"
    "Y286LjNmfSIsICJtwrMiXSwKICAgICAgICAgWyJBZ3VhIiwgZiJ7cHJvcF9hZ3VhXzI6LjJmfcOXe1BFU09fU0FDT19DRU1F"
    "TlRPfSIsIGYie2FndWFfc2FjbzouMWZ9IiwgIkx0Il1dLAogICAgICAgIGNvbHVtbnM9WyJNYXRlcmlhbCIsICJDw6FsY3Vs"
    "byIsICJQcm9wb3JjacOzbiIsICJVbmlkYWQiXSkKICAgIHNlYygiUEFTTyAxNMKwOiBQUk9QT1JDSU9ORVMgRU4gVk9MVU1F"
    "TiAobcKzKSIsCiAgICAgICAgIlNlIGNvbnZpZXJ0ZSBsYSBwcm9wb3JjacOzbiBlbiBwZXNvIGEgdm9sdW1lbiB1c2FuZG8g"
    "ZWwgUGVzbyBVbml0YXJpbyAiCiAgICAgICAgIlN1ZWx0byBTZWNvIChQVVNTKSBkZSBjYWRhIGFncmVnYWRvLiBFbCBjZW1l"
    "bnRvIHNlIHRvbWEgY29tbyByZWZlcmVuY2lhOiAiCiAgICAgICAgZiIxIHNhY28gPSB7UEVTT19TQUNPX0NFTUVOVE99IEtn"
    "LiIsCiAgICAgICAgZiJGw7NybXVsYSBnZW5lcmFsOiBWb2wgKG3CsykgPSAoUHJvcG9yY2nDs24gw5cge1BFU09fU0FDT19D"
    "RU1FTlRPfSBLZy9zYWNvKSAvIFBVU1MgKEtnL23CsykiLAogICAgICAgICJDZW1lbnRvOiAxIHNhY28gKHJlZmVyZW5jaWEg"
    "ZmlqYSkiLAogICAgICAgIGYiQWcuIEZpbm86ICh7cHJvcF9maW5vXzI6LjJmfSDDlyB7UEVTT19TQUNPX0NFTUVOVE99KSAv"
    "IHtkWydhcmVuYV9wdXNzJ106LjBmfSA9ICIKICAgICAgICBmIntwcm9wX2Zpbm9fMiAqIFBFU09fU0FDT19DRU1FTlRPOi4z"
    "Zn0gLyB7ZFsnYXJlbmFfcHVzcyddOi4wZn0gPSB7dm9sX2Zpbm9fc2FjbzouM2Z9IG3CsyIsCiAgICAgICAgZiJBZy4gR3J1"
    "ZXNvOiAoe3Byb3BfZ3J1ZXNvXzI6LjJmfSDDlyB7UEVTT19TQUNPX0NFTUVOVE99KSAvIHtkWydwaWVkcmFfcHVzcyddOi4w"
    "Zn0gPSAiCiAgICAgICAgZiJ7cHJvcF9ncnVlc29fMiAqIFBFU09fU0FDT19DRU1FTlRPOi4zZn0gLyB7ZFsncGllZHJhX3B1"
    "c3MnXTouMGZ9ID0ge3ZvbF9ncnVlc29fc2FjbzouM2Z9IG3CsyIsCiAgICAgICAgZiJBZ3VhOiB7cHJvcF9hZ3VhXzI6LjJm"
    "fSDDlyB7UEVTT19TQUNPX0NFTUVOVE99ID0ge2FndWFfc2FjbzouMWZ9IEx0IHBvciBzYWNvIGRlIGNlbWVudG8iLAogICAg"
    "ICAgIHRhYmxhX3ZvbF9zYWNvLAogICAgICAgIGYiUHJvcG9yY2nDs24gZW4gdm9sdW1lbjogMSBzYWNvIDoge3ZvbF9maW5v"
    "X3NhY286LjNmfSBtwrMgOiB7dm9sX2dydWVzb19zYWNvOi4zZn0gbcKzICIKICAgICAgICBmImNvbiB7YWd1YV9zYWNvOi4x"
    "Zn0gTHQgZGUgYWd1YSIsCiAgICAgICAgIkxlY3R1cmEgZW4gb2JyYTogcG9yIGNhZGEgc2FjbyBkZSBjZW1lbnRvIHNlIHVz"
    "YW4gIgogICAgICAgIGYie3ZvbF9maW5vX3NhY286LjNmfSBtwrMgZGUgYXJlbmEgeSB7dm9sX2dydWVzb19zYWNvOi4zZn0g"
    "bcKzIGRlIHBpZWRyYSwgbWVkaWRvcyBlbiBzdWVsdG8uIikKCiAgICBEX20gPSBkWyJwcm9iZXRhX2QiXSAvIDEwMC4wCiAg"
    "ICBIX20gPSBkWyJwcm9iZXRhX2giXSAvIDEwMC4wCiAgICB2X3Byb2JldGEgPSAobnAucGkgLyA0LjApICogRF9tICoqIDIg"
    "KiBIX20KICAgIHZfcHJvYmV0YXMgPSBkWyJuX3Byb2JldGFzIl0gKiB2X3Byb2JldGEKICAgIHZfdG90YWwgPSByb3VuZCh2"
    "X3Byb2JldGFzICogKDEgKyBkWyJkZXNwZXJkaWNpbyJdIC8gMTAwLjApLCA0KQoKICAgIGNhbnRfY2VtZW50byA9IGNlbWVu"
    "dG9faHVtZWRvICogdl90b3RhbAogICAgY2FudF9maW5vID0gZmlub19odW1lZG9fcmVkICogdl90b3RhbAogICAgY2FudF9n"
    "cnVlc28gPSBncnVlc29faHVtZWRvX3JlZCAqIHZfdG90YWwKICAgIGNhbnRfYWd1YSA9IGFndWFfZWZlY3RpdmEgKiB2X3Rv"
    "dGFsCiAgICBudW1fc2Fjb3MgPSBjYW50X2NlbWVudG8gLyBQRVNPX1NBQ09fQ0VNRU5UTwogICAgb2JyYV9maW5vID0gdm9s"
    "X2Zpbm9fc2FjbyAqIHZfdG90YWwKICAgIG9icmFfZ3J1ZXNvID0gdm9sX2dydWVzb19zYWNvICogdl90b3RhbAogICAgb2Jy"
    "YV9hZ3VhID0gYWd1YV9zYWNvICogdl90b3RhbAoKICAgIFIudXBkYXRlKGRpY3Qodl9wcm9iZXRhPXZfcHJvYmV0YSwgdl9w"
    "cm9iZXRhcz12X3Byb2JldGFzLCB2X3RvdGFsPXZfdG90YWwsCiAgICAgICAgICAgICAgICAgIGNhbnRfY2VtZW50bz1jYW50"
    "X2NlbWVudG8sIGNhbnRfZmlubz1jYW50X2Zpbm8sIGNhbnRfZ3J1ZXNvPWNhbnRfZ3J1ZXNvLAogICAgICAgICAgICAgICAg"
    "ICBjYW50X2FndWE9Y2FudF9hZ3VhLCBudW1fc2Fjb3M9bnVtX3NhY29zLCBvYnJhX2Zpbm89b2JyYV9maW5vLAogICAgICAg"
    "ICAgICAgICAgICBvYnJhX2dydWVzbz1vYnJhX2dydWVzbywgb2JyYV9hZ3VhPW9icmFfYWd1YSkpCgogICAgdGFibGFfY2Fu"
    "dCA9IHBkLkRhdGFGcmFtZSgKICAgICAgICBbWyJDZW1lbnRvIChow7ptZWRvKSIsIGYie2NlbWVudG9faHVtZWRvOi4yZn0g"
    "S2cvbcKzIiwgZiLDlyB7dl90b3RhbDouM2Z9IiwKICAgICAgICAgIGYie2NhbnRfY2VtZW50bzouM2Z9IEtnIl0sCiAgICAg"
    "ICAgIFsiQWcuIEZpbm8gKGjDum1lZG8pIiwgZiJ7Zmlub19odW1lZG9fcmVkOi4yZn0gS2cvbcKzIiwgZiLDlyB7dl90b3Rh"
    "bDouM2Z9IiwKICAgICAgICAgIGYie2NhbnRfZmlubzouM2Z9IEtnIl0sCiAgICAgICAgIFsiQWcuIEdydWVzbyAoaMO6bWVk"
    "bykiLCBmIntncnVlc29faHVtZWRvX3JlZDouMWZ9IEtnL23CsyIsIGYiw5cge3ZfdG90YWw6LjNmfSIsCiAgICAgICAgICBm"
    "IntjYW50X2dydWVzbzouM2Z9IEtnIl0sCiAgICAgICAgIFsiQWd1YSBlZmVjdGl2YSIsIGYie2FndWFfZWZlY3RpdmE6LjJm"
    "fSBMdC9twrMiLCBmIsOXIHt2X3RvdGFsOi4zZn0iLAogICAgICAgICAgZiJ7Y2FudF9hZ3VhOi4zZn0gTHQiXV0sCiAgICAg"
    "ICAgY29sdW1ucz1bIk1hdGVyaWFsIiwgIkRvc2lzIChLZyBvIEx0IC8gbcKzKSIsIGYiw5cgViA9IHt2X3RvdGFsOi4zZn0g"
    "bcKzIiwgIkNhbnRpZGFkIl0pCgogICAgdGFibGFfb2JyYSA9IHBkLkRhdGFGcmFtZSgKICAgICAgICBbWyJDZW1lbnRvIiwg"
    "ZiJ7Y2FudF9jZW1lbnRvOi4zZn0gw7cge1BFU09fU0FDT19DRU1FTlRPfSIsIGYie251bV9zYWNvczouM2Z9IiwgInNhY29z"
    "Il0sCiAgICAgICAgIFsiQWcuIEZpbm8iLCBmInt2b2xfZmlub19zYWNvOi4zZn0gw5cge3ZfdG90YWw6LjNmfSIsIGYie29i"
    "cmFfZmlubzouM2Z9IiwgIm3CsyJdLAogICAgICAgICBbIkFnLiBHcnVlc28iLCBmInt2b2xfZ3J1ZXNvX3NhY286LjNmfSDD"
    "lyB7dl90b3RhbDouM2Z9IiwgZiJ7b2JyYV9ncnVlc286LjNmfSIsICJtwrMiXSwKICAgICAgICAgWyJBZ3VhIGVmZWN0aXZh"
    "IiwgZiJ7YWd1YV9zYWNvOi4xZn0gw5cge3ZfdG90YWw6LjNmfSIsIGYie29icmFfYWd1YTouM2Z9IiwgIkx0Il1dLAogICAg"
    "ICAgIGNvbHVtbnM9WyJNYXRlcmlhbCIsICJDw6FsY3VsbyIsICJDYW50aWRhZCIsICJVbmlkYWQiXSkKCiAgICBzZWMoZiJQ"
    "QVNPIDE1wrA6IERPU0lGSUNBQ0nDk04gUEFSQSBQUk9CRVRBUyBERSAiCiAgICAgICAgZiJ7ZFsncHJvYmV0YV9kJ106LjBm"
    "fcOXe2RbJ3Byb2JldGFfaCddOi4wZn0gY20iLAogICAgICAgICJEYXRvczoiLAogICAgICAgIGYi4oCiIERpbWVuc2lvbmVz"
    "IGRlIGxhIHByb2JldGE6IGRpw6FtZXRybyA9IHtkWydwcm9iZXRhX2QnXTouMGZ9IGNtLCAiCiAgICAgICAgZiJhbHR1cmEg"
    "PSB7ZFsncHJvYmV0YV9oJ106LjBmfSBjbSIsCiAgICAgICAgZiLigKIgTsO6bWVybyBkZSBwcm9iZXRhczoge2RbJ25fcHJv"
    "YmV0YXMnXX0iLAogICAgICAgIGYi4oCiIERlc3BlcmRpY2lvOiB7ZFsnZGVzcGVyZGljaW8nXTouMGZ9JSIsCiAgICAgICAg"
    "IjEuIFZvbHVtZW4gZGUgdW5hIHByb2JldGEgY2lsw61uZHJpY2E6IiwKICAgICAgICBmIlYgPSDPgC80IMOXIETCsiDDlyBI"
    "ID0gMC43ODUgw5cgKHtEX206LjJmfSnCsiDDlyB7SF9tOi4yZn0gbSIsCiAgICAgICAgZiJWID0ge3ZfcHJvYmV0YTouM2Z9"
    "IG3Csy9wcm9iZXRhIiwKICAgICAgICBmIjIuIFZvbHVtZW4gdG90YWwgcGFyYSB7ZFsnbl9wcm9iZXRhcyddfSBwcm9iZXRh"
    "czoiLAogICAgICAgIGYiVjMgPSB7dl9wcm9iZXRhczouM2Z9IG3CsyIsCiAgICAgICAgZiIzLiBWb2x1bWVuIGNvbiBkZXNw"
    "ZXJkaWNpbyBkZWwge2RbJ2Rlc3BlcmRpY2lvJ106LjBmfSU6IiwKICAgICAgICBmIlZ0b3RhbCA9IHt2X3Byb2JldGFzOi4z"
    "Zn0gw5cgKDEgKyB7ZFsnZGVzcGVyZGljaW8nXTouMGZ9LzEwMCkgPSAiCiAgICAgICAgZiJ7dl9wcm9iZXRhczouM2Z9IMOX"
    "IHsxICsgZFsnZGVzcGVyZGljaW8nXSAvIDEwMDouMmZ9ID0ge3ZfdG90YWw6LjNmfSBtwrMiLAogICAgICAgIGYiVnRvdGFs"
    "IOKJiCB7dl90b3RhbDouM2Z9IG3CsyIsCiAgICAgICAgIjQuIENhbnRpZGFkIGRlIGNhZGEgbWF0ZXJpYWwgPSBEb3NpcyBw"
    "b3IgbcKzIMOXIFZ0b3RhbDoiLAogICAgICAgIHRhYmxhX2NhbnQsCiAgICAgICAgIjUuIEVxdWl2YWxlbnRlIGVuIHZvbHVt"
    "ZW4gcG9yIHNhY28gKHByb3BvcmNpb25lcyBkZSBvYnJhKToiLAogICAgICAgIGYiTsO6bWVybyBkZSBzYWNvcyDiiYgge2Nh"
    "bnRfY2VtZW50bzouM2Z9IEtnIMO3IHtQRVNPX1NBQ09fQ0VNRU5UT30gS2cvc2FjbyA9ICIKICAgICAgICBmIntudW1fc2Fj"
    "b3M6LjNmfSBzYWNvcyIsCiAgICAgICAgZiJBZy4gRmlubyA9IHtvYnJhX2Zpbm86LjNmfSBtwrMiLAogICAgICAgIGYiQWcu"
    "IEdydWVzbyA9IHtvYnJhX2dydWVzbzouM2Z9IG3CsyIsCiAgICAgICAgZiJBZ3VhID0ge29icmFfYWd1YTouM2Z9IEx0IiwK"
    "ICAgICAgICB0YWJsYV9vYnJhKQoKICAgIHRhYmxhX3Jlc3VtZW4gPSBwZC5EYXRhRnJhbWUoCiAgICAgICAgW1siQ2VtZW50"
    "byIsIGYie2NhbnRfY2VtZW50bzouM2Z9IEtnIiwgZiJ7bnVtX3NhY29zOi4zZn0iLCAic2Fjb3MiXSwKICAgICAgICAgWyJB"
    "Zy4gRmlubyIsIGYie2NhbnRfZmlubzouM2Z9IEtnIiwgZiJ7b2JyYV9maW5vOi4zZn0iLCAibcKzIl0sCiAgICAgICAgIFsi"
    "QWcuIEdydWVzbyIsIGYie2NhbnRfZ3J1ZXNvOi4zZn0gS2ciLCBmIntvYnJhX2dydWVzbzouM2Z9IiwgIm3CsyJdLAogICAg"
    "ICAgICBbIkFndWEgZWZlY3RpdmEiLCBmIntjYW50X2FndWE6LjNmfSBMdCIsIGYie29icmFfYWd1YTouM2Z9IiwgIkx0Il1d"
    "LAogICAgICAgIGNvbHVtbnM9WyJNYXRlcmlhbCIsICJQZXNvIGjDum1lZG8iLCAiVm9sdW1lbiBvYnJhIiwgIlVuaWRhZCB2"
    "b2wuIl0pCiAgICBzZWMoZiJSRVNVTUVOIOKAlCBEb3NpZmljYWNpw7NuIHBhcmEge2RbJ25fcHJvYmV0YXMnXX0gcHJvYmV0"
    "YXMgIgogICAgICAgIGYiKHtkWydwcm9iZXRhX2QnXTouMGZ9w5d7ZFsncHJvYmV0YV9oJ106LjBmfSBjbSkgKyB7ZFsnZGVz"
    "cGVyZGljaW8nXTouMGZ9JSBkZXNwZXJkaWNpbzoiLAogICAgICAgIHRhYmxhX3Jlc3VtZW4sCiAgICAgICAgIlZPTFVNRU4g"
    "VE9UQUwiLAogICAgICAgIGYiViA9IHt2X3RvdGFsOi4zZn0gbcKzIikKCiAgICByZXR1cm4gUiwgUwoKCmRlZiByZW5kZXJf"
    "c2VjY2lvbmVzKGQsIFIsIFMpOgogICAgTCA9IFtdCiAgICBMLmFwcGVuZCgiPSIgKiBBTkNITykKICAgIEwuYXBwZW5kKCJU"
    "RUNOT0xPR0lBIERFTCBDT05DUkVUTyIuY2VudGVyKEFOQ0hPKSkKICAgIEwuYXBwZW5kKFJbInRpdHVsbyJdLmNlbnRlcihB"
    "TkNITykpCiAgICBMLmFwcGVuZCgiPSIgKiBBTkNITykKICAgIEwuYXBwZW5kKGYiRmVjaGEgICAgICAgICA6IHtkYXRldGlt"
    "ZS5ub3coKS5zdHJmdGltZSgnJVktJW0tJWQgJUg6JU06JVMnKX0iKQogICAgTC5hcHBlbmQoZiJSZXNwb25zYWJsZSAgIDog"
    "e2RbJ3Jlc3BvbnNhYmxlJ119IikKICAgIEwuYXBwZW5kKGYiUmVzaXN0ZW5jaWEgICA6IGYnYyA9IHtkWydmYyddOi4wZn0g"
    "a2cvY20yICAgICIKICAgICAgICAgICAgIGYiZidjciA9IHtSWydmY3InXTouMGZ9IGtnL2NtMiIpCiAgICBMLmFwcGVuZChm"
    "IlJlbGFjaW9uIGEvYyAgOiB7UlsnYWMnXTouMmZ9IikKICAgIEwuYXBwZW5kKCIiKQoKICAgIGZvciBzIGluIFM6CiAgICAg"
    "ICAgTC5hcHBlbmQoIi0iICogQU5DSE8pCiAgICAgICAgTC5hcHBlbmQoc1sidGl0dWxvIl0pCiAgICAgICAgTC5hcHBlbmQo"
    "Ii0iICogQU5DSE8pCiAgICAgICAgZm9yIHRpcG8sIGNvbnRlbmlkbyBpbiBzWyJibG9xdWVzIl06CiAgICAgICAgICAgIGlm"
    "IHRpcG8gPT0gInRhYmxhIjoKICAgICAgICAgICAgICAgIEwuYXBwZW5kKGNvbnRlbmlkby50b19zdHJpbmcoaW5kZXg9RmFs"
    "c2UpKQogICAgICAgICAgICBlbGlmIGNvbnRlbmlkbzoKICAgICAgICAgICAgICAgIEwuYXBwZW5kKHN0cihjb250ZW5pZG8p"
    "KQogICAgICAgIEwuYXBwZW5kKCIiKQoKICAgIEwuYXBwZW5kKCI9IiAqIEFOQ0hPKQogICAgTC5hcHBlbmQoIlBST1BPUkNJ"
    "T04gRklOQUwgRU4gT0JSQSIuY2VudGVyKEFOQ0hPKSkKICAgIEwuYXBwZW5kKCI9IiAqIEFOQ0hPKQogICAgTC5hcHBlbmQo"
    "ZiIgIDEgc2FjbyBkZSBjZW1lbnRvICh7UEVTT19TQUNPX0NFTUVOVE99IEtnKSA6ICIKICAgICAgICAgICAgIGYie1JbJ3Zv"
    "bF9maW5vX3NhY28nXTouM2Z9IG0zIGRlIGFyZW5hIDoge1JbJ3ZvbF9ncnVlc29fc2FjbyddOi4zZn0gbTMgZGUgcGllZHJh"
    "IikKICAgIEwuYXBwZW5kKGYiICBBZ3VhOiB7UlsnYWd1YV9zYWNvJ106LjFmfSBsdCBwb3Igc2FjbyAgICAoYS9jID0ge1Jb"
    "J2FjJ106LjJmfSkiKQogICAgTC5hcHBlbmQoZiIgIENlbWVudG86IHtSWydjZW1lbnRvX2h1bWVkbyddOi4yZn0ga2cvbTMg"
    "ICBBZ3VhIGVmZWN0aXZhOiB7UlsnYWd1YV9lZmVjdGl2YSddOi4yZn0gbHQvbTMiKQogICAgTC5hcHBlbmQoIj0iICogQU5D"
    "SE8pCiAgICBMLmFwcGVuZChkWyJyZXNwb25zYWJsZSJdLmNlbnRlcihBTkNITykpCiAgICBMLmFwcGVuZCgiPSIgKiBBTkNI"
    "TykKICAgIHJldHVybiBMCgpkZWYgaW1wcmltaXJfcmVwb3J0ZShsaW5lYXMpOgogICAgZm9yIGwgaW4gbGluZWFzOgogICAg"
    "ICAgIHByaW50KGwpCgpDX1RJVFVMTyA9ICIxRjM4NjQiCkNfU0VDQ0lPTiA9ICIyRTc1QjYiCkNfQ0FCRUNFUkEgPSAiRDlF"
    "MUYyIgpDX0NMQVZFID0gIkZDRTRENiIKQ19UT1RBTCA9ICJGRkYyQ0MiCkNfREFUTyA9ICJGMkYyRjIiCkNfUkVTVUxUQURP"
    "ID0gIkM2RTBCNCIKCmRlZiBfYWx0b19uZWNlc2FyaW8odGV4dG8sIGFuY2hvKToKICAgIGFuY2hvID0gbWF4KDYsIGludChh"
    "bmNobyAqIDAuOTUpKQogICAgbGluZWFzID0gMAogICAgZm9yIHBhcnRlIGluIHN0cih0ZXh0bykuc3BsaXQoIlxuIik6CiAg"
    "ICAgICAgbGluZWFzICs9IG1heCgxLCAobGVuKHBhcnRlKSArIGFuY2hvIC0gMSkgLy8gYW5jaG8pCiAgICByZXR1cm4gbWF4"
    "KDEsIGxpbmVhcykKCk1BUkNBU19URVhUTyA9ICgKICAgICJQcm9wb3JjacOzbiBlbiBwZXNvOiIsICJQcm9wb3JjacOzbiBl"
    "biB2b2x1bWVuOiIsICJWdG90YWwg4omIIiwKICAgICJWT0xVTUVOIFRPVEFMIiwgIk7Dum1lcm8gZGUgc2Fjb3MiLAopCk1B"
    "UkNBU19OVU1FUk8gPSAoCiAgICAiRidjciA9ICIsICJSIGEvYyA9ICIsICJiL2JvID0gIiwgIkFpcmUgQXRyYXBhZG8gPSAi"
    "LCAiQWd1YSA9ICIsCiAgICAiUGVzbyBBZy4gR3J1ZXNvID0gIiwgIlZvbC4gQWJzb2x1dG8gQWcuIEZpbm8gPSAiLCAiUGVz"
    "byBTZWNvIEFnLiBGaW5vID0gIiwKICAgICJcdTAzYTMgQWd1YSBMaWJyZSA9ICIsICJBZ3VhIEVmZWN0aXZhID0gIiwgIkMg"
    "PSAiLCAiViA9ICIsCikKX1BBVFJPTkVTX1JFU1VMVEFETyA9IHR1cGxlKHJlLmNvbXBpbGUocmUuZXNjYXBlKG0pICsgciJc"
    "cyotP1xkIikgZm9yIG0gaW4gTUFSQ0FTX05VTUVSTykKX1BBVFJPTl9DX01JTiA9IHJlLmNvbXBpbGUociIoPzwhW1x3J10p"
    "YyA9IFxzKi0/XGQiKQoKZGVmIF9lc19yZXN1bHRhZG8odGV4dG8pOgogICAgdCA9ICIgIi5qb2luKHN0cih0ZXh0bykuc3Bs"
    "aXQoKSkKICAgIGlmIGFueShtYXJjYSBpbiB0IGZvciBtYXJjYSBpbiBNQVJDQVNfVEVYVE8pOgogICAgICAgIHJldHVybiBU"
    "cnVlCiAgICBpZiBfUEFUUk9OX0NfTUlOLnNlYXJjaCh0KToKICAgICAgICByZXR1cm4gVHJ1ZQogICAgcmV0dXJuIGFueShw"
    "YXRyb24uc2VhcmNoKHQpIGZvciBwYXRyb24gaW4gX1BBVFJPTkVTX1JFU1VMVEFETykKCmRlZiBfYm9yZGUoKToKICAgIGxh"
    "ZG8gPSBTaWRlKHN0eWxlPSJ0aGluIiwgY29sb3I9IkIwQjBCMCIpCiAgICByZXR1cm4gQm9yZGVyKGxlZnQ9bGFkbywgcmln"
    "aHQ9bGFkbywgdG9wPWxhZG8sIGJvdHRvbT1sYWRvKQoKZGVmIF9lc2NyaWJpcl90aXR1bG8od3MsIGZpbGEsIHRleHRvLCBu"
    "Y29scyk6CiAgICB3cy5tZXJnZV9jZWxscyhzdGFydF9yb3c9ZmlsYSwgc3RhcnRfY29sdW1uPTEsIGVuZF9yb3c9ZmlsYSwg"
    "ZW5kX2NvbHVtbj1uY29scykKICAgIGMgPSB3cy5jZWxsKHJvdz1maWxhLCBjb2x1bW49MSwgdmFsdWU9dGV4dG8pCiAgICBj"
    "LmZvbnQgPSBGb250KGJvbGQ9VHJ1ZSwgc2l6ZT0xNCwgY29sb3I9IkZGRkZGRiIpCiAgICBjLmZpbGwgPSBQYXR0ZXJuRmls"
    "bCgic29saWQiLCBmZ0NvbG9yPUNfVElUVUxPKQogICAgYy5hbGlnbm1lbnQgPSBBbGlnbm1lbnQoaG9yaXpvbnRhbD0iY2Vu"
    "dGVyIiwgdmVydGljYWw9ImNlbnRlciIsIHdyYXBfdGV4dD1UcnVlKQogICAgYW5jaG9fdG90YWwgPSBzdW0oKHdzLmNvbHVt"
    "bl9kaW1lbnNpb25zW2dldF9jb2x1bW5fbGV0dGVyKGkpXS53aWR0aCBvciA4LjQzKQogICAgICAgICAgICAgICAgICAgICAg"
    "Zm9yIGkgaW4gcmFuZ2UoMSwgbmNvbHMgKyAxKSkKICAgIHdzLnJvd19kaW1lbnNpb25zW2ZpbGFdLmhlaWdodCA9IG1heCgy"
    "Ni4wLCBfYWx0b19uZWNlc2FyaW8odGV4dG8sIGFuY2hvX3RvdGFsKSAqIDE2LjAgKyA2KQogICAgcmV0dXJuIGZpbGEgKyAx"
    "CgpkZWYgX2VzY3JpYmlyX3NlY2Npb24od3MsIGZpbGEsIHRleHRvLCBuY29scyk6CiAgICB3cy5tZXJnZV9jZWxscyhzdGFy"
    "dF9yb3c9ZmlsYSwgc3RhcnRfY29sdW1uPTEsIGVuZF9yb3c9ZmlsYSwgZW5kX2NvbHVtbj1uY29scykKICAgIGMgPSB3cy5j"
    "ZWxsKHJvdz1maWxhLCBjb2x1bW49MSwgdmFsdWU9dGV4dG8pCiAgICBjLmZvbnQgPSBGb250KGJvbGQ9VHJ1ZSwgc2l6ZT0x"
    "MSwgY29sb3I9IkZGRkZGRiIpCiAgICBjLmZpbGwgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBmZ0NvbG9yPUNfU0VDQ0lPTikK"
    "ICAgIGMuYWxpZ25tZW50ID0gQWxpZ25tZW50KGhvcml6b250YWw9ImxlZnQiLCB2ZXJ0aWNhbD0iY2VudGVyIiwgaW5kZW50"
    "PTEsIHdyYXBfdGV4dD1UcnVlKQogICAgYW5jaG9fdG90YWwgPSBzdW0oKHdzLmNvbHVtbl9kaW1lbnNpb25zW2dldF9jb2x1"
    "bW5fbGV0dGVyKGkpXS53aWR0aCBvciA4LjQzKQogICAgICAgICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UoMSwgbmNv"
    "bHMgKyAxKSkKICAgIHdzLnJvd19kaW1lbnNpb25zW2ZpbGFdLmhlaWdodCA9IG1heCgxOS4wLCBfYWx0b19uZWNlc2FyaW8o"
    "dGV4dG8sIGFuY2hvX3RvdGFsKSAqIDE0LjAgKyA0KQogICAgcmV0dXJuIGZpbGEgKyAxCgpkZWYgX2VzY3JpYmlyX2NhYmVj"
    "ZXJhKHdzLCBmaWxhLCBlbmNhYmV6YWRvcyk6CiAgICBtYXhfbGluZWFzID0gMQogICAgZm9yIGksIHRleHRvIGluIGVudW1l"
    "cmF0ZShlbmNhYmV6YWRvcywgc3RhcnQ9MSk6CiAgICAgICAgYyA9IHdzLmNlbGwocm93PWZpbGEsIGNvbHVtbj1pLCB2YWx1"
    "ZT10ZXh0bykKICAgICAgICBjLmZvbnQgPSBGb250KGJvbGQ9VHJ1ZSwgY29sb3I9Q19USVRVTE8pCiAgICAgICAgYy5maWxs"
    "ID0gUGF0dGVybkZpbGwoInNvbGlkIiwgZmdDb2xvcj1DX0NBQkVDRVJBKQogICAgICAgIGMuYWxpZ25tZW50ID0gQWxpZ25t"
    "ZW50KGhvcml6b250YWw9ImNlbnRlciIsIHZlcnRpY2FsPSJjZW50ZXIiLCB3cmFwX3RleHQ9VHJ1ZSkKICAgICAgICBjLmJv"
    "cmRlciA9IF9ib3JkZSgpCiAgICAgICAgYW5jaG8gPSB3cy5jb2x1bW5fZGltZW5zaW9uc1tnZXRfY29sdW1uX2xldHRlcihp"
    "KV0ud2lkdGggb3IgOC40MwogICAgICAgIG1heF9saW5lYXMgPSBtYXgobWF4X2xpbmVhcywgX2FsdG9fbmVjZXNhcmlvKHRl"
    "eHRvLCBhbmNobykpCiAgICB3cy5yb3dfZGltZW5zaW9uc1tmaWxhXS5oZWlnaHQgPSBtYXgoMjAuMCwgbWF4X2xpbmVhcyAq"
    "IDE0LjAgKyA1KQogICAgcmV0dXJuIGZpbGEgKyAxCgpkZWYgX2VzY3JpYmlyX2ZpbGEod3MsIGZpbGEsIHZhbG9yZXMsIG5l"
    "Z3JpdGE9RmFsc2UsIHJlbGxlbm89Tm9uZSwgY2VudHJhcl9kZXNkZT0yKToKICAgIG1heF9saW5lYXMgPSAxCiAgICBmb3Ig"
    "aSwgdiBpbiBlbnVtZXJhdGUodmFsb3Jlcywgc3RhcnQ9MSk6CiAgICAgICAgYyA9IHdzLmNlbGwocm93PWZpbGEsIGNvbHVt"
    "bj1pLCB2YWx1ZT12KQogICAgICAgIGMuYm9yZGVyID0gX2JvcmRlKCkKICAgICAgICBjLmZvbnQgPSBGb250KGJvbGQ9bmVn"
    "cml0YSkKICAgICAgICBpZiByZWxsZW5vOgogICAgICAgICAgICBjLmZpbGwgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBmZ0Nv"
    "bG9yPXJlbGxlbm8pCiAgICAgICAgaWYgaSA+PSBjZW50cmFyX2Rlc2RlOgogICAgICAgICAgICBjLmFsaWdubWVudCA9IEFs"
    "aWdubWVudChob3Jpem9udGFsPSJjZW50ZXIiLCB2ZXJ0aWNhbD0iY2VudGVyIikKICAgICAgICBlbHNlOgogICAgICAgICAg"
    "ICBjLmFsaWdubWVudCA9IEFsaWdubWVudCh2ZXJ0aWNhbD0iY2VudGVyIiwgd3JhcF90ZXh0PVRydWUpCiAgICAgICAgYW5j"
    "aG8gPSB3cy5jb2x1bW5fZGltZW5zaW9uc1tnZXRfY29sdW1uX2xldHRlcihpKV0ud2lkdGggb3IgOC40MwogICAgICAgIG1h"
    "eF9saW5lYXMgPSBtYXgobWF4X2xpbmVhcywgX2FsdG9fbmVjZXNhcmlvKHYsIGFuY2hvKSkKICAgIGlmIG1heF9saW5lYXMg"
    "PiAxOgogICAgICAgIHdzLnJvd19kaW1lbnNpb25zW2ZpbGFdLmhlaWdodCA9IG1heF9saW5lYXMgKiAxNC4wICsgMgogICAg"
    "cmV0dXJuIGZpbGEgKyAxCgpkZWYgX2VzY3JpYmlyX25vdGEod3MsIGZpbGEsIHRleHRvLCBuY29scyk6CiAgICBpbXBvcnRh"
    "bnRlID0gX2VzX3Jlc3VsdGFkbyh0ZXh0bykKICAgIGlmIGltcG9ydGFudGU6CiAgICAgICAgZm9yIGkgaW4gcmFuZ2UoMSwg"
    "bmNvbHMgKyAxKToKICAgICAgICAgICAgd3MuY2VsbChyb3c9ZmlsYSwgY29sdW1uPWkpLmZpbGwgPSBQYXR0ZXJuRmlsbCgi"
    "c29saWQiLCBmZ0NvbG9yPUNfUkVTVUxUQURPKQogICAgd3MubWVyZ2VfY2VsbHMoc3RhcnRfcm93PWZpbGEsIHN0YXJ0X2Nv"
    "bHVtbj0xLCBlbmRfcm93PWZpbGEsIGVuZF9jb2x1bW49bmNvbHMpCiAgICBjID0gd3MuY2VsbChyb3c9ZmlsYSwgY29sdW1u"
    "PTEsIHZhbHVlPXRleHRvKQogICAgaWYgaW1wb3J0YW50ZToKICAgICAgICBjLmZvbnQgPSBGb250KGJvbGQ9VHJ1ZSwgc2l6"
    "ZT0xMCwgY29sb3I9IjFGMzg2NCIpCiAgICBlbHNlOgogICAgICAgIGMuZm9udCA9IEZvbnQoaXRhbGljPVRydWUsIHNpemU9"
    "MTAsIGNvbG9yPSI1OTU5NTkiKQogICAgYy5hbGlnbm1lbnQgPSBBbGlnbm1lbnQodmVydGljYWw9ImNlbnRlciIsIHdyYXBf"
    "dGV4dD1UcnVlLCBpbmRlbnQ9MSkKICAgIGFuY2hvX3RvdGFsID0gc3VtKCh3cy5jb2x1bW5fZGltZW5zaW9uc1tnZXRfY29s"
    "dW1uX2xldHRlcihpKV0ud2lkdGggb3IgOC40MykKICAgICAgICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKDEsIG5j"
    "b2xzICsgMSkpCiAgICB3cy5yb3dfZGltZW5zaW9uc1tmaWxhXS5oZWlnaHQgPSBtYXgoMTUuMCwgX2FsdG9fbmVjZXNhcmlv"
    "KHRleHRvLCBhbmNob190b3RhbCkgKiAxNC4wKQogICAgcmV0dXJuIGZpbGEgKyAxCgpkZWYgX2FqdXN0YXJfYW5jaG9zKHdz"
    "LCBhbmNob3MpOgogICAgZm9yIGksIGFuY2hvIGluIGVudW1lcmF0ZShhbmNob3MsIHN0YXJ0PTEpOgogICAgICAgIHdzLmNv"
    "bHVtbl9kaW1lbnNpb25zW2dldF9jb2x1bW5fbGV0dGVyKGkpXS53aWR0aCA9IGFuY2hvCgpkZWYgX2VzY3JpYmlyX3NlY2Np"
    "b25fY29tcGxldGEod3MsIGZpbGEsIHNlY2Npb24sIG5jb2xzKToKICAgIGZpbGEgPSBfZXNjcmliaXJfc2VjY2lvbih3cywg"
    "ZmlsYSwgc2VjY2lvblsidGl0dWxvIl0sIG5jb2xzKQogICAgZm9yIHRpcG8sIGNvbnRlbmlkbyBpbiBzZWNjaW9uWyJibG9x"
    "dWVzIl06CiAgICAgICAgaWYgdGlwbyA9PSAidGFibGEiOgogICAgICAgICAgICB0YWJsYSA9IGNvbnRlbmlkbwogICAgICAg"
    "ICAgICBmaWxhID0gX2VzY3JpYmlyX2NhYmVjZXJhKHdzLCBmaWxhLCBbc3RyKGMpIGZvciBjIGluIHRhYmxhLmNvbHVtbnNd"
    "KQogICAgICAgICAgICBmb3IgXywgcmVnaXN0cm8gaW4gdGFibGEuaXRlcnJvd3MoKToKICAgICAgICAgICAgICAgIGZpbGEg"
    "PSBfZXNjcmliaXJfZmlsYSh3cywgZmlsYSwgW3N0cih2KSBmb3IgdiBpbiByZWdpc3Ryby52YWx1ZXNdKQogICAgICAgIGVs"
    "aWYgY29udGVuaWRvOgogICAgICAgICAgICBmaWxhID0gX2VzY3JpYmlyX25vdGEod3MsIGZpbGEsIHN0cihjb250ZW5pZG8p"
    "LCBuY29scykKICAgIHJldHVybiBmaWxhCgpkZWYgZXhwb3J0YXJfZXhjZWwoZCwgUiwgUywgcnV0YSk6CiAgICB3YiA9IFdv"
    "cmtib29rKCkKICAgIHdiLnJlbW92ZSh3Yi5hY3RpdmUpCgogICAgd3MgPSB3Yi5jcmVhdGVfc2hlZXQoIkRvc2lmaWNhY2lv"
    "biBGaW5hbCIpCiAgICBfYWp1c3Rhcl9hbmNob3Mod3MsIFsyOCwgMTYsIDE4LCAxOCwgMTJdKQogICAgZiA9IDEKICAgIGYg"
    "PSBfZXNjcmliaXJfdGl0dWxvKHdzLCBmLCAiRE9TSUZJQ0FDSU9OIERFIENPTkNSRVRPIC0gTUVUT0RPIEFDSSAyMTEiLCA1"
    "KQogICAgZiA9IF9lc2NyaWJpcl9maWxhKHdzLCBmLCBbIlByb3llY3RvIiwgZFsiZGVzY3JpcGNpb24iXSwgIiIsCiAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAiRmVjaGEiLCBkYXRldGltZS5ub3coKS5zdHJmdGltZSgiJVktJW0tJWQgJUg6"
    "JU0iKV0pCiAgICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYsIFsiUmVzcG9uc2FibGUiLCBkWyJyZXNwb25zYWJsZSJdLCAi"
    "IiwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICJDb25zaXN0ZW5jaWEiLCBkWyJjb25zaXN0ZW5jaWEiXV0pCiAg"
    "ICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYsIFsiUmVzaXN0ZW5jaWEgZidjIChrZy9jbTIpIiwgZiJ7ZFsnZmMnXTouMGZ9"
    "IiwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICJSZXNpc3RlbmNpYSBmJ2NyIChrZy9jbTIpIiwgZiJ7UlsnZmNy"
    "J106LjBmfSIsCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBmImEvYyA9IHtSWydhYyddOi4yZn0iXSkKICAgIGYg"
    "Kz0gMQoKICAgIGYgPSBfZXNjcmliaXJfc2VjY2lvbih3cywgZiwgIkRPU0lGSUNBQ0lPTiBQQVJBIEVMIFZPTFVNRU4gQ0FM"
    "Q1VMQURPIiwgNSkKICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiTUFURVJJQUwiLCAiRE9TSVMgUE9SIG3C"
    "syIsICJQRVNPIChrZykiLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICJWT0xVTUVOIERFIE9CUkEiLCAi"
    "VU5JREFEIl0pCiAgICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYsIFsiQ2VtZW50byIsIGYie1JbJ2NlbWVudG9faHVtZWRv"
    "J106LjJmfSIsCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBmIntSWydjYW50X2NlbWVudG8nXTouM2Z9IiwgZiJ7"
    "UlsnbnVtX3NhY29zJ106LjNmfSIsICJzYWNvcyJdLAogICAgICAgICAgICAgICAgICAgICAgIHJlbGxlbm89Q19EQVRPKQog"
    "ICAgZiA9IF9lc2NyaWJpcl9maWxhKHdzLCBmLCBbIkFncmVnYWRvIGZpbm8gKGFyZW5hKSIsIGYie1JbJ2Zpbm9faHVtZWRv"
    "X3JlZCddOi4yZn0iLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZiJ7UlsnY2FudF9maW5vJ106LjNmfSIsIGYi"
    "e1JbJ29icmFfZmlubyddOi4zZn0iLCAibcKzIl0sCiAgICAgICAgICAgICAgICAgICAgICAgcmVsbGVubz1DX0RBVE8pCiAg"
    "ICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYsIFsiQWdyZWdhZG8gZ3J1ZXNvIChwaWVkcmEpIiwgZiJ7UlsnZ3J1ZXNvX2h1"
    "bWVkb19yZWQnXTouMWZ9IiwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGYie1JbJ2NhbnRfZ3J1ZXNvJ106LjNm"
    "fSIsIGYie1JbJ29icmFfZ3J1ZXNvJ106LjNmfSIsICJtwrMiXSwKICAgICAgICAgICAgICAgICAgICAgICByZWxsZW5vPUNf"
    "REFUTykKICAgIGYgPSBfZXNjcmliaXJfZmlsYSh3cywgZiwgWyJBZ3VhIGVmZWN0aXZhIiwgZiJ7UlsnYWd1YV9lZmVjdGl2"
    "YSddOi4yZn0iLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZiJ7UlsnY2FudF9hZ3VhJ106LjNmfSIsIGYie1Jb"
    "J29icmFfYWd1YSddOi4zZn0iLCAiTHQiXSwKICAgICAgICAgICAgICAgICAgICAgICByZWxsZW5vPUNfREFUTykKICAgIGYg"
    "PSBfZXNjcmliaXJfZmlsYSh3cywgZiwgWyJWT0xVTUVOIFRPVEFMIiwgIiIsICIiLCBmIntSWyd2X3RvdGFsJ106LjNmfSIs"
    "ICJtwrMiXSwKICAgICAgICAgICAgICAgICAgICAgICBuZWdyaXRhPVRydWUsIHJlbGxlbm89Q19UT1RBTCkKICAgIGYgKz0g"
    "MQoKICAgIGZvciBzIGluIFM6CiAgICAgICAgaWYgc1sidGl0dWxvIl0uc3RhcnRzd2l0aCgoIlBBU08gMTMiLCAiUEFTTyAx"
    "NCIsICJQQVNPIDE1IiwgIlJFU1VNRU4iKSk6CiAgICAgICAgICAgIGYgPSBfZXNjcmliaXJfc2VjY2lvbl9jb21wbGV0YSh3"
    "cywgZiwgcywgNSkKICAgICAgICAgICAgZiArPSAxCgogICAgd3MgPSB3Yi5jcmVhdGVfc2hlZXQoIkRhdG9zIGRlIEVudHJh"
    "ZGEiKQogICAgX2FqdXN0YXJfYW5jaG9zKHdzLCBbNDAsIDIwLCAxNF0pCiAgICBmID0gMQogICAgZiA9IF9lc2NyaWJpcl90"
    "aXR1bG8od3MsIGYsICJEQVRPUyBERSBFTlRSQURBIC0gUFJPWUVDVE8iLCAzKQoKICAgIGRlZiBibG9xdWUodGl0dWxvLCBm"
    "aWxhcyk6CiAgICAgICAgbm9ubG9jYWwgZgogICAgICAgIGYgPSBfZXNjcmliaXJfc2VjY2lvbih3cywgZiwgdGl0dWxvLCAz"
    "KQogICAgICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiUEFSQU1FVFJPIiwgIlZBTE9SIiwgIlVOSURBRCJd"
    "KQogICAgICAgIGZvciBpLCAocGFyYW0sIHZhbG9yLCB1bmlkYWQpIGluIGVudW1lcmF0ZShmaWxhcyk6CiAgICAgICAgICAg"
    "IGYgPSBfZXNjcmliaXJfZmlsYSh3cywgZiwgW3BhcmFtLCB2YWxvciwgdW5pZGFkXSwKICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIHJlbGxlbm89KENfREFUTyBpZiBpICUgMiA9PSAwIGVsc2UgTm9uZSkpCgogICAgYmxvcXVlKCJEQVRPUyBE"
    "RUwgRElTRU5PIiwgWwogICAgICAgICgiRGVzY3JpcGNpb24gZGVsIGVsZW1lbnRvIiwgZFsiZGVzY3JpcGNpb24iXSwgIi0i"
    "KSwKICAgICAgICAoIlJlc2lzdGVuY2lhIGVzcGVjaWZpY2FkYSBmJ2MiLCBmIntkWydmYyddOi4wZn0iLCAia2cvY20yIiks"
    "CiAgICAgICAgKCJSZXNpc3RlbmNpYSBwcm9tZWRpbyBmJ2NyIiwgZiJ7UlsnZmNyJ106LjBmfSIsICJrZy9jbTIiKSwKICAg"
    "ICAgICAoIkNvbnNpc3RlbmNpYSIsIGRbImNvbnNpc3RlbmNpYSJdLCAiLSIpLAogICAgICAgICgiQXNlbnRhbWllbnRvIiwg"
    "Ulsic2x1bXBfcmFuZ28iXSwgIi0iKSwKICAgICAgICAoIkNvbiBhaXJlIGluY29ycG9yYWRvIiwgIlNpIiBpZiBkWyJjb25f"
    "YWlyZSJdIGVsc2UgIk5vIiwgIi0iKSwKICAgICAgICAoIkV4cG9zaWNpb24iLCBkWyJleHBvc2ljaW9uX25vbWJyZSJdLCAi"
    "LSIpLAogICAgXSkKICAgIGYgKz0gMQogICAgYmxvcXVlKCJEQVRPUyBERUwgQ0VNRU5UTyIsIFsKICAgICAgICAoIlRpcG8g"
    "ZGUgY2VtZW50byIsIGRbImNlbWVudG9fbm9tYnJlIl0sICItIiksCiAgICAgICAgKCJQZXNvIGVzcGVjaWZpY28iLCBmIntk"
    "WydwZV9jZW1lbnRvJ106LjJmfSIsICJnL2NtMyIpLAogICAgICAgICgiRG9zaXMgZGUgY2VtZW50byAoZGlzZW5vKSIsIGYi"
    "e1JbJ2NlbWVudG9fZGlzZW5vJ106LjJmfSIsICJrZy9tMyIpLAogICAgICAgICgiRG9zaXMgZGUgY2VtZW50byAoaHVtZWRv"
    "IGNvcnJlZ2lkbykiLCBmIntSWydjZW1lbnRvX2h1bWVkbyddOi4yZn0iLCAia2cvbTMiKSwKICAgIF0pCiAgICBmICs9IDEK"
    "ICAgIGJsb3F1ZSgiREFUT1MgREVMIEFHUkVHQURPIEZJTk8gKEFSRU5BKSIsIFsKICAgICAgICAoIlBlc28gdW5pdGFyaW8g"
    "c3VlbHRvIHNlY28gKFBVU1MpIiwgZiJ7ZFsnYXJlbmFfcHVzcyddOi4wZn0iLCAia2cvbTMiKSwKICAgICAgICAoIlBlc28g"
    "dW5pdGFyaW8gY29tcGFjdGFkbyBzZWNvIChQVUNTKSIsIGYie2RbJ2FyZW5hX3B1Y3MnXTouMGZ9IiwgImtnL20zIiksCiAg"
    "ICAgICAgKCJQZXNvIGVzcGVjaWZpY28gZGUgbWFzYSIsIGYie2RbJ2FyZW5hX3BlJ106LjJmfSIsICItIiksCiAgICAgICAg"
    "KCJDb250ZW5pZG8gZGUgaHVtZWRhZCIsIGYie2RbJ2FyZW5hX2h1bWVkYWQnXTouMWZ9IiwgIiUiKSwKICAgICAgICAoIlBv"
    "cmNlbnRhamUgZGUgYWJzb3JjaW9uIiwgZiJ7ZFsnYXJlbmFfYWJzb3JjaW9uJ106LjFmfSIsICIlIiksCiAgICAgICAgKCJN"
    "b2R1bG8gZGUgZmluZXphIiwgZiJ7ZFsnYXJlbmFfbWYnXTouMmZ9IiwgIi0iKSwKICAgICAgICAoIlBlc28gc2VjbyBlbiBs"
    "YSBtZXpjbGEiLCBmIntSWydwZXNvX2Zpbm8nXTouM2Z9IiwgImtnL20zIiksCiAgICAgICAgKCJQZXNvIGh1bWVkbyBjb3Jy"
    "ZWdpZG8iLCBmIntSWydmaW5vX2h1bWVkb19yZWQnXTouMGZ9IiwgImtnL20zIiksCiAgICBdKQogICAgZiArPSAxCiAgICBi"
    "bG9xdWUoIkRBVE9TIERFTCBBR1JFR0FETyBHUlVFU08gKFBJRURSQSkiLCBbCiAgICAgICAgKCJQZXNvIHVuaXRhcmlvIHN1"
    "ZWx0byBzZWNvIChQVVNTKSIsIGYie2RbJ3BpZWRyYV9wdXNzJ106LjBmfSIsICJrZy9tMyIpLAogICAgICAgICgiUGVzbyB1"
    "bml0YXJpbyBjb21wYWN0YWRvIHNlY28gKFBVQ1MpIiwgZiJ7ZFsncGllZHJhX3B1Y3MnXTouMGZ9IiwgImtnL20zIiksCiAg"
    "ICAgICAgKCJQZXNvIGVzcGVjaWZpY28gZGUgbWFzYSIsIGYie2RbJ3BpZWRyYV9wZSddOi4yZn0iLCAiLSIpLAogICAgICAg"
    "ICgiQ29udGVuaWRvIGRlIGh1bWVkYWQiLCBmIntkWydwaWVkcmFfaHVtZWRhZCddOi4xZn0iLCAiJSIpLAogICAgICAgICgi"
    "UG9yY2VudGFqZSBkZSBhYnNvcmNpb24iLCBmIntkWydwaWVkcmFfYWJzb3JjaW9uJ106LjFmfSIsICIlIiksCiAgICAgICAg"
    "KCJUYW1hbm8gbWF4aW1vIG5vbWluYWwgKFRNTikiLCBkWyJ0bW4iXSwgIi0iKSwKICAgICAgICAoIlZvbHVtZW4gZGUgYWdy"
    "ZWdhZG8gZ3J1ZXNvIChiL2JvKSIsIGYie1JbJ2JibyddOi4yZn0iLCAiLSIpLAogICAgICAgICgiUGVzbyBzZWNvIGVuIGxh"
    "IG1lemNsYSIsIGYie1JbJ3Blc29fZ3J1ZXNvJ106LjFmfSIsICJrZy9tMyIpLAogICAgICAgICgiUGVzbyBodW1lZG8gY29y"
    "cmVnaWRvIiwgZiJ7UlsnZ3J1ZXNvX2h1bWVkb19yZWQnXTouMGZ9IiwgImtnL20zIiksCiAgICBdKQogICAgZiArPSAxCiAg"
    "ICBibG9xdWUoIkRBVE9TIERFTCBBR1VBIiwgWwogICAgICAgICgiQWd1YSBkZSBtZXpjbGFkbyAoZGlzZW5vKSIsIGYie1Jb"
    "J2FndWEnXTouMGZ9IiwgImx0L20zIiksCiAgICAgICAgKCJBZ3VhIGxpYnJlIGRlbCBhZ3JlZ2FkbyBmaW5vIiwgZiJ7Ulsn"
    "YWd1YV9saWJyZV9maW5vJ106LjJmfSIsICJsdC9tMyIpLAogICAgICAgICgiQWd1YSBsaWJyZSBkZWwgYWdyZWdhZG8gZ3J1"
    "ZXNvIiwgZiIre1JbJ2FndWFfbGlicmVfZ3J1ZXNvJ106LjJmfSIsICJsdC9tMyIpLAogICAgICAgICgiQWd1YSBsaWJyZSB0"
    "b3RhbCIsIGYie1JbJ2FndWFfbGlicmVfdG90YWwnXTouMmZ9IiwgImx0L20zIiksCiAgICAgICAgKCJBZ3VhIGVmZWN0aXZh"
    "IiwgZiJ7UlsnYWd1YV9lZmVjdGl2YSddOi4yZn0iLCAibHQvbTMiKSwKICAgICAgICAoIlJlbGFjaW9uIGFndWEvY2VtZW50"
    "byAoYS9jKSIsIGYie1JbJ2FjJ106LjJmfSIsICItIiksCiAgICAgICAgKCJBZ3VhIHBvciBzYWNvIGRlIGNlbWVudG8iLCBm"
    "IntSWydhZ3VhX3NhY28nXTouMWZ9IiwgImx0L3NhY28iKSwKICAgIF0pCiAgICBmICs9IDEKICAgIGJsb3F1ZSgiREFUT1Mg"
    "REUgTEFTIFBST0JFVEFTIiwgWwogICAgICAgICgiRGlhbWV0cm8iLCBmIntkWydwcm9iZXRhX2QnXTouMGZ9IiwgImNtIiks"
    "CiAgICAgICAgKCJBbHR1cmEiLCBmIntkWydwcm9iZXRhX2gnXTouMGZ9IiwgImNtIiksCiAgICAgICAgKCJOdW1lcm8gZGUg"
    "cHJvYmV0YXMiLCBmIntkWyduX3Byb2JldGFzJ119IiwgIi0iKSwKICAgICAgICAoIkRlc3BlcmRpY2lvIiwgZiJ7ZFsnZGVz"
    "cGVyZGljaW8nXTouMGZ9IiwgIiUiKSwKICAgICAgICAoIlZvbHVtZW4gdG90YWwgYSBwcmVwYXJhciIsIGYie1JbJ3ZfdG90"
    "YWwnXTouM2Z9IiwgIm0zIiksCiAgICBdKQoKICAgIHdzID0gd2IuY3JlYXRlX3NoZWV0KCJQYXNvIGEgUGFzbyIpCiAgICBf"
    "YWp1c3Rhcl9hbmNob3Mod3MsIFszNCwgMTYsIDE2LCAxNiwgMTYsIDE2LCAxNiwgMTYsIDE2XSkKICAgIGYgPSAxCiAgICBm"
    "ID0gX2VzY3JpYmlyX3RpdHVsbyh3cywgZiwgIkRJU0XDkU8gREUgTUVaQ0xBUyDigJQgTcOJVE9ETyBDT01JVMOJIEFDSSAy"
    "MTEiLCA5KQogICAgZm9yIHMgaW4gUzoKICAgICAgICBmID0gX2VzY3JpYmlyX3NlY2Npb25fY29tcGxldGEod3MsIGYsIHMs"
    "IDkpCiAgICAgICAgZiArPSAxCgogICAgd3MgPSB3Yi5jcmVhdGVfc2hlZXQoIlRhYmxhcyBBQ0kiKQogICAgX2FqdXN0YXJf"
    "YW5jaG9zKHdzLCBbMTYsIDE0LCAxNCwgMTQsIDE0LCAxNCwgMTQsIDE0LCAxNF0pCiAgICBmID0gMQogICAgZiA9IF9lc2Ny"
    "aWJpcl90aXR1bG8od3MsIGYsICJUQUJMQVMgQUNJIDIxMSAtIFZBTE9SRVMgREUgUkVGRVJFTkNJQSIsIDkpCgogICAgZiA9"
    "IF9lc2NyaWJpcl9zZWNjaW9uKHdzLCBmLCAiVEFCTEEgMDEgLSBWT0xVTUVOIFVOSVRBUklPIERFIEFHVUEgKGx0L20zKSIs"
    "IDkpCiAgICBmb3IgdGlwbywgZXRpcXVldGEgaW4gKCgic2luX2FpcmUiLCAiQ09OQ1JFVE8gU0lOIEFJUkUgSU5DT1JQT1JB"
    "RE8iKSwKICAgICAgICAgICAgICAgICAgICAgICAgICAgKCJjb25fYWlyZSIsICJDT05DUkVUTyBDT04gQUlSRSBJTkNPUlBP"
    "UkFETyIpKToKICAgICAgICBmID0gX2VzY3JpYmlyX25vdGEod3MsIGYsIGV0aXF1ZXRhLCA5KQogICAgICAgIGYgPSBfZXNj"
    "cmliaXJfY2FiZWNlcmEod3MsIGYsIFsiQXNlbnRhbWllbnRvIl0gKyBUTU5fRElTRU5PKQogICAgICAgIGZvciByIGluIFsn"
    "MSIgYSAyIicsICczIiBhIDQiJywgJzYiIGEgNyInXToKICAgICAgICAgICAgZiA9IF9lc2NyaWJpcl9maWxhKHdzLCBmLCBb"
    "cl0gKyBbZm9ybWF0ZWFyKFRBQkxBXzAxX0FHVUFbdGlwb11bcl1bdF0sIDApCiAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgIGZvciB0IGluIFRNTl9ESVNFTk9dKQogICAgZiArPSAxCgogICAgZiA9IF9lc2NyaWJpcl9z"
    "ZWNjaW9uKHdzLCBmLCAiVEFCTEEgMDQgLSBQRVNPIERFTCBBR1JFR0FETyBHUlVFU08gKGIvYm8pIiwgOSkKICAgIGYgPSBf"
    "ZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiVE1OIl0gKyBbZiJNRiB7bTouMmZ9IiBmb3IgbSBpbiAoMi40MCwgMi42MCwg"
    "Mi44MCwgMy4wMCldKQogICAgZm9yIHQgaW4gVE1OX0RJU0VOTzoKICAgICAgICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYs"
    "IFt0XSArIFtmIntUQUJMQV8wNF9BR19HUlVFU09bdF1bbV06LjJmfSIKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICBmb3IgbSBpbiAoMi40MCwgMi42MCwgMi44MCwgMy4wMCldKQogICAgZiArPSAxCgogICAgZiA9IF9lc2Ny"
    "aWJpcl9zZWNjaW9uKHdzLCBmLCAiVEFCTEEgMDUgLSBSRUxBQ0lPTiBBR1VBL0NFTUVOVE8gUE9SIFJFU0lTVEVOQ0lBIiwg"
    "OSkKICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiZidjciAoa2cvY20yKSIsICJTaW4gYWlyZSIsICJDb24g"
    "YWlyZSJdKQogICAgZm9yIGZjXywgdiBpbiBzb3J0ZWQoVEFCTEFfMDVfQUMuaXRlbXMoKSk6CiAgICAgICAgZiA9IF9lc2Ny"
    "aWJpcl9maWxhKHdzLCBmLCBbZmNfLCBmb3JtYXRlYXIodlswXSwgMiksIGZvcm1hdGVhcih2WzFdLCAyKV0pCiAgICBmICs9"
    "IDEKCiAgICBmID0gX2VzY3JpYmlyX3NlY2Npb24od3MsIGYsICJUQUJMQSAwNiAtIENPTlRFTklETyBERSBBSVJFIElOQ09S"
    "UE9SQURPIFkgVE9UQUwgKCUpIiwgOSkKICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiVE1OIiwgIlN1YXZl"
    "IiwgIk1vZGVyYWRhIiwgIlNldmVyYSJdKQogICAgZm9yIHQsIHYgaW4gVEFCTEFfMDZfQUlSRV9JTkNPUlBPUkFETy5pdGVt"
    "cygpOgogICAgICAgIGYgPSBfZXNjcmliaXJfZmlsYSh3cywgZiwgW3QsIGYie3ZbMF06LjFmfSIsIGYie3ZbMV06LjFmfSIs"
    "IGYie3ZbMl06LjFmfSJdKQoKICAgIHdzID0gd2IuY3JlYXRlX3NoZWV0KCJHcmFmaWNvcyIpCiAgICBfYWp1c3Rhcl9hbmNo"
    "b3Mod3MsIFszMCwgMTgsIDE0LCAxNCwgMTQsIDE0LCAxNCwgMTQsIDE0XSkKICAgIGYgPSAxCiAgICBmID0gX2VzY3JpYmly"
    "X3RpdHVsbyh3cywgZiwgIkdSQUZJQ09TIERFIExBIERPU0lGSUNBQ0lPTiIsIDkpCiAgICBmID0gX2VzY3JpYmlyX25vdGEo"
    "d3MsIGYsICJMb3MgZ3JhZmljb3Mgc2UgZ2VuZXJhbiBjb24gbG9zIHJlc3VsdGFkb3MgZGVsIGRpc2Vuby4iLCA5KQogICAg"
    "ZiArPSAxCgogICAgZiA9IF9lc2NyaWJpcl9zZWNjaW9uKHdzLCBmLCAiRE9TSVMgUE9SIG0zIChLZyBvIEx0KSIsIDkpCiAg"
    "ICBmaWxhX2Rvc2lzID0gZiArIDEKICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiTWF0ZXJpYWwiLCAiRG9z"
    "aXMgcG9yIG0zIl0pCiAgICBmb3Igbm9tLCB2YWwgaW4gKCgiQ2VtZW50byIsIFJbImNlbWVudG9faHVtZWRvIl0pLAogICAg"
    "ICAgICAgICAgICAgICAgICAoIkFnLiBGaW5vIiwgUlsiZmlub19odW1lZG9fcmVkIl0pLAogICAgICAgICAgICAgICAgICAg"
    "ICAoIkFnLiBHcnVlc28iLCBSWyJncnVlc29faHVtZWRvX3JlZCJdKSwKICAgICAgICAgICAgICAgICAgICAgKCJBZ3VhIiwg"
    "UlsiYWd1YV9lZmVjdGl2YSJdKSk6CiAgICAgICAgZiA9IF9lc2NyaWJpcl9maWxhKHdzLCBmLCBbbm9tLCByb3VuZChmbG9h"
    "dCh2YWwpLCAyKV0sIHJlbGxlbm89Q19EQVRPKQogICAgZmluX2Rvc2lzID0gZiAtIDEKICAgIGYgKz0gMQoKICAgIGYgPSBf"
    "ZXNjcmliaXJfc2VjY2lvbih3cywgZiwgIlBST1BPUkNJT04gRU4gVk9MVU1FTiBQT1IgU0FDTyAobTMpIiwgOSkKICAgIGZp"
    "bGFfdm9sID0gZiArIDEKICAgIGYgPSBfZXNjcmliaXJfY2FiZWNlcmEod3MsIGYsIFsiTWF0ZXJpYWwiLCAiVm9sdW1lbiAo"
    "bTMpIl0pCiAgICBmb3Igbm9tLCB2YWwgaW4gKCgiQWcuIEZpbm8iLCByb3VuZChSWyJ2b2xfZmlub19zYWNvIl0sIDMpKSwK"
    "ICAgICAgICAgICAgICAgICAgICAgKCJBZy4gR3J1ZXNvIiwgcm91bmQoUlsidm9sX2dydWVzb19zYWNvIl0sIDMpKSk6CiAg"
    "ICAgICAgZiA9IF9lc2NyaWJpcl9maWxhKHdzLCBmLCBbbm9tLCB2YWxdLCByZWxsZW5vPUNfREFUTykKICAgIGZpbl92b2wg"
    "PSBmIC0gMQogICAgZiArPSAxCgogICAgZiA9IF9lc2NyaWJpcl9zZWNjaW9uKHdzLCBmLCAiQ0FOVElEQURFUyBQQVJBIExB"
    "UyBQUk9CRVRBUyIsIDkpCiAgICBmaWxhX3Byb2IgPSBmICsgMQogICAgZiA9IF9lc2NyaWJpcl9jYWJlY2VyYSh3cywgZiwg"
    "WyJNYXRlcmlhbCIsICJDYW50aWRhZCAoS2cgbyBMdCkiXSkKICAgIGZvciBub20sIHZhbCBpbiAoKCJDZW1lbnRvIiwgUlsi"
    "Y2FudF9jZW1lbnRvIl0pLAogICAgICAgICAgICAgICAgICAgICAoIkFnLiBGaW5vIiwgUlsiY2FudF9maW5vIl0pLAogICAg"
    "ICAgICAgICAgICAgICAgICAoIkFnLiBHcnVlc28iLCBSWyJjYW50X2dydWVzbyJdKSwKICAgICAgICAgICAgICAgICAgICAg"
    "KCJBZ3VhIiwgUlsiY2FudF9hZ3VhIl0pKToKICAgICAgICBmID0gX2VzY3JpYmlyX2ZpbGEod3MsIGYsIFtub20sIHJvdW5k"
    "KGZsb2F0KHZhbCksIDMpXSwgcmVsbGVubz1DX0RBVE8pCiAgICBmaW5fcHJvYiA9IGYgLSAxCgogICAgZzEgPSBCYXJDaGFy"
    "dCgpCiAgICBnMS50eXBlID0gImNvbCIKICAgIGcxLnN0eWxlID0gMTAKICAgIGcxLnRpdGxlID0gIkRvc2lzIHBvciBtMyBk"
    "ZSBjb25jcmV0byIKICAgIGcxLnlfYXhpcy50aXRsZSA9ICJLZyBvIEx0IgogICAgZzEuYWRkX2RhdGEoUmVmZXJlbmNlKHdz"
    "LCBtaW5fY29sPTIsIG1pbl9yb3c9ZmlsYV9kb3NpcyAtIDEsIG1heF9yb3c9ZmluX2Rvc2lzKSwKICAgICAgICAgICAgICAg"
    "IHRpdGxlc19mcm9tX2RhdGE9VHJ1ZSkKICAgIGcxLnNldF9jYXRlZ29yaWVzKFJlZmVyZW5jZSh3cywgbWluX2NvbD0xLCBt"
    "aW5fcm93PWZpbGFfZG9zaXMsIG1heF9yb3c9ZmluX2Rvc2lzKSkKICAgIGcxLmhlaWdodCA9IDguNQogICAgZzEud2lkdGgg"
    "PSAxNQogICAgd3MuYWRkX2NoYXJ0KGcxLCAiRDIiKQoKICAgIGcyID0gUGllQ2hhcnQoKQogICAgZzIudGl0bGUgPSAiUHJv"
    "cG9yY2lvbiBlbiB2b2x1bWVuIHBvciBzYWNvIChtMykiCiAgICBnMi5hZGRfZGF0YShSZWZlcmVuY2Uod3MsIG1pbl9jb2w9"
    "MiwgbWluX3Jvdz1maWxhX3ZvbCAtIDEsIG1heF9yb3c9ZmluX3ZvbCksCiAgICAgICAgICAgICAgICB0aXRsZXNfZnJvbV9k"
    "YXRhPVRydWUpCiAgICBnMi5zZXRfY2F0ZWdvcmllcyhSZWZlcmVuY2Uod3MsIG1pbl9jb2w9MSwgbWluX3Jvdz1maWxhX3Zv"
    "bCwgbWF4X3Jvdz1maW5fdm9sKSkKICAgIGcyLmhlaWdodCA9IDguNQogICAgZzIud2lkdGggPSAxMwogICAgd3MuYWRkX2No"
    "YXJ0KGcyLCAiRDE5IikKCiAgICBnMyA9IEJhckNoYXJ0KCkKICAgIGczLnR5cGUgPSAiYmFyIgogICAgZzMuc3R5bGUgPSAx"
    "MgogICAgZzMudGl0bGUgPSAiQ2FudGlkYWRlcyBwYXJhIGxhcyBwcm9iZXRhcyIKICAgIGczLnlfYXhpcy50aXRsZSA9ICJL"
    "ZyBvIEx0IgogICAgZzMuYWRkX2RhdGEoUmVmZXJlbmNlKHdzLCBtaW5fY29sPTIsIG1pbl9yb3c9ZmlsYV9wcm9iIC0gMSwg"
    "bWF4X3Jvdz1maW5fcHJvYiksCiAgICAgICAgICAgICAgICB0aXRsZXNfZnJvbV9kYXRhPVRydWUpCiAgICBnMy5zZXRfY2F0"
    "ZWdvcmllcyhSZWZlcmVuY2Uod3MsIG1pbl9jb2w9MSwgbWluX3Jvdz1maWxhX3Byb2IsIG1heF9yb3c9ZmluX3Byb2IpKQog"
    "ICAgZzMuaGVpZ2h0ID0gOC41CiAgICBnMy53aWR0aCA9IDE1CiAgICB3cy5hZGRfY2hhcnQoZzMsICJEMzYiKQoKICAgIG9y"
    "ZGVuID0gWyJQYXNvIGEgUGFzbyIsICJEb3NpZmljYWNpb24gRmluYWwiLCAiR3JhZmljb3MiLAogICAgICAgICAgICAgIkRh"
    "dG9zIGRlIEVudHJhZGEiLCAiVGFibGFzIEFDSSJdCiAgICB3Yi5fc2hlZXRzID0gW3diW25vbWJyZV0gZm9yIG5vbWJyZSBp"
    "biBvcmRlbl0KCiAgICB3Yi5zYXZlKHJ1dGEpCgpkZWYgbW9zdHJhcl9yZXN1bWVuX2ZpbmFsKGQsIFIpOgogICAgc2Fjb3Mg"
    "PSBSWyJjZW1lbnRvX2h1bWVkbyJdIC8gUEVTT19TQUNPX0NFTUVOVE8KCiAgICBlbmNhYmV6YWRvKCJSRVNVTUVOIERFIExB"
    "IERPU0lGSUNBQ0lPTiAtIEFDSSAyMTEiKQoKICAgIHByaW50KGYiICBQcm95ZWN0byAgICAgIDoge2RbJ2Rlc2NyaXBjaW9u"
    "J119IikKICAgIHByaW50KGYiICBSZXNwb25zYWJsZSAgIDoge2RbJ3Jlc3BvbnNhYmxlJ119IikKICAgIHByaW50KGYiICBG"
    "ZWNoYSAgICAgICAgIDoge2RhdGV0aW1lLm5vdygpLnN0cmZ0aW1lKCclWS0lbS0lZCAlSDolTScpfSIpCiAgICBsaW5lYSgi"
    "LSIpCiAgICBwcmludChmIiAgUmVzaXN0ZW5jaWEgZidjICAgOiB7ZFsnZmMnXTouMGZ9IGtnL2NtMiIpCiAgICBwcmludChm"
    "IiAgUmVzaXN0ZW5jaWEgZidjciAgOiB7UlsnZmNyJ106LjBmfSBrZy9jbTIiKQogICAgcHJpbnQoZiIgIFJlbGFjaW9uIGEv"
    "YyAgICAgIDoge1JbJ2FjJ106LjJmfSIpCiAgICBwcmludChmIiAgQ29uc2lzdGVuY2lhICAgICAgOiB7ZFsnY29uc2lzdGVu"
    "Y2lhJ119ICAoe1JbJ3NsdW1wX3JhbmdvJ119KSIpCgogICAgcHJpbnQoKQogICAgbGluZWEoIi0iKQogICAgcHJpbnQoIiAg"
    "RE9TSVMgUE9SIG0zIERFIENPTkNSRVRPIChwZXNvIGh1bWVkbyBjb3JyZWdpZG8pIikKICAgIGxpbmVhKCItIikKICAgIHBy"
    "aW50KGYiICAgIENlbWVudG8gLi4uLi4uLi4uLi4uLiB7UlsnY2VtZW50b19odW1lZG8nXTo4LjJmfSBrZy9tMyAgICh7c2Fj"
    "b3M6LjJmfSBzYWNvcykiKQogICAgcHJpbnQoZiIgICAgQWdyZWdhZG8gZmlubyAuLi4uLi4uIHtSWydmaW5vX2h1bWVkb19y"
    "ZWQnXTo4LjJmfSBrZy9tMyIpCiAgICBwcmludChmIiAgICBBZ3JlZ2FkbyBncnVlc28gLi4uLi4ge1JbJ2dydWVzb19odW1l"
    "ZG9fcmVkJ106OC4xZn0ga2cvbTMiKQogICAgcHJpbnQoZiIgICAgQWd1YSBlZmVjdGl2YSAuLi4uLi4uIHtSWydhZ3VhX2Vm"
    "ZWN0aXZhJ106OC4yZn0gbHQvbTMiKQoKICAgIHByaW50KCkKICAgIGxpbmVhKCItIikKICAgIHByaW50KCIgIFBST1BPUkNJ"
    "T04gRU4gUEVTTyAgKGNlbWVudG8gOiBmaW5vIDogZ3J1ZXNvKSIpCiAgICBsaW5lYSgiLSIpCiAgICBwcmludChmIiAgICAx"
    "IDoge1JbJ3Byb3BfZmlubyddOi4yZn0gOiB7UlsncHJvcF9ncnVlc28nXTouMmZ9ICAgICAgYS9jID0ge1JbJ2FjJ106LjJm"
    "fSIpCgogICAgcHJpbnQoKQogICAgbGluZWEoIi0iKQogICAgcHJpbnQoZiIgIFBST1BPUkNJT04gRU4gVk9MVU1FTiBQT1Ig"
    "U0FDTyBERSBDRU1FTlRPICh7UEVTT19TQUNPX0NFTUVOVE99IGtnKSIpCiAgICBsaW5lYSgiLSIpCiAgICBwcmludChmIiAg"
    "ICBBZ3JlZ2FkbyBmaW5vIC4uLi4uLi4ge1JbJ3ZvbF9maW5vX3NhY28nXTouM2Z9IG0zIikKICAgIHByaW50KGYiICAgIEFn"
    "cmVnYWRvIGdydWVzbyAuLi4uLiB7Ulsndm9sX2dydWVzb19zYWNvJ106LjNmfSBtMyIpCiAgICBwcmludChmIiAgICBBZ3Vh"
    "IC4uLi4uLi4uLi4uLi4uLi4ge1JbJ2FndWFfc2FjbyddOi4xZn0gbHQiKQoKICAgIHByaW50KCkKICAgIGxpbmVhKCItIikK"
    "ICAgIHByaW50KGYiICBDQU5USURBREVTIFBBUkEgRUwgVk9MVU1FTiBDQUxDVUxBRE8gKFYgdG90YWwgPSB7Ulsndl90b3Rh"
    "bCddOi4zZn0gbTMpIikKICAgIGxpbmVhKCItIikKICAgIHByaW50KGYiICAgIENlbWVudG8gLi4uLi4uLi4uLi4uLiB7Ulsn"
    "Y2FudF9jZW1lbnRvJ106OC4zZn0ga2cgICAoe1JbJ251bV9zYWNvcyddOi4zZn0gc2Fjb3MpIikKICAgIHByaW50KGYiICAg"
    "IEFncmVnYWRvIGZpbm8gLi4uLi4uLiB7UlsnY2FudF9maW5vJ106OC4zZn0ga2ciKQogICAgcHJpbnQoZiIgICAgQWdyZWdh"
    "ZG8gZ3J1ZXNvIC4uLi4uIHtSWydjYW50X2dydWVzbyddOjguM2Z9IGtnIikKICAgIHByaW50KGYiICAgIEFndWEgZWZlY3Rp"
    "dmEgLi4uLi4uLiB7UlsnY2FudF9hZ3VhJ106OC4zZn0gbHQiKQogICAgbGluZWEoIj0iKQoKZGVmIG1haW4oKToKICAgIGNh"
    "cnBldGEgPSBvcy5wYXRoLmRpcm5hbWUob3MucGF0aC5hYnNwYXRoKF9fZmlsZV9fKSkKCiAgICB3aGlsZSBUcnVlOgogICAg"
    "ICAgIGxpbXBpYXJfcGFudGFsbGEoKQogICAgICAgIGVuY2FiZXphZG8oIlRFQ05PTE9HSUEgREVMIENPTkNSRVRPIC0gRElT"
    "RU5PIERFIE1FWkNMQSBBQ0kgMjExIikKICAgICAgICBwcmludCgiICAxKSBEaXNlbm8gZGUgbWV6Y2xhIEFDSSAyMTEiKQog"
    "ICAgICAgIHByaW50KCIgIDIpIFNhbGlyIikKICAgICAgICBsaW5lYSgiPSIpCiAgICAgICAgb3BjaW9uID0gaW5wdXQoIlNl"
    "bGVjY2lvbmUgdW5hIG9wY2lvbjogIikuc3RyaXAoKQoKICAgICAgICBpZiBvcGNpb24gPT0gIjEiOgogICAgICAgICAgICBs"
    "aW1waWFyX3BhbnRhbGxhKCkKICAgICAgICAgICAgZCA9IHNvbGljaXRhcl9kYXRvcygpCgogICAgICAgICAgICBsaW1waWFy"
    "X3BhbnRhbGxhKCkKICAgICAgICAgICAgbW9zdHJhcl9yZXN1bWVuX2RhdG9zKGQpCiAgICAgICAgICAgIGlmIG5vdCBwZWRp"
    "cl9zaV9ubygiXG5Mb3MgZGF0b3Mgc29uIGNvcnJlY3Rvcz8gKHMvbikgW3NdOiAiLCAicyIpOgogICAgICAgICAgICAgICAg"
    "cGF1c2EoIlxuVm9sdmllbmRvIGFsIG1lbnUuLi4gUHJlc2lvbmUgRU5URVIuIikKICAgICAgICAgICAgICAgIGNvbnRpbnVl"
    "CgogICAgICAgICAgICBsaW1waWFyX3BhbnRhbGxhKCkKICAgICAgICAgICAgcHJpbnQoIlxuQ2FsY3VsYW5kbyBkaXNlbm8g"
    "ZGUgbWV6Y2xhLi4uIikKICAgICAgICAgICAgUiwgUyA9IGRpc2VuYXIoZCkKCiAgICAgICAgICAgIGxpbXBpYXJfcGFudGFs"
    "bGEoKQogICAgICAgICAgICBtb3N0cmFyX3Jlc3VtZW5fZmluYWwoZCwgUikKCiAgICAgICAgICAgIHByaW50KCJcbkRlc2Vh"
    "IGhhY2VyIGVsIGluZm9ybWUgZGUgbGEgZG9zaWZpY2FjaW9uPyIpCiAgICAgICAgICAgIGhhY2VyX2luZm9ybWUgPSBwZWRp"
    "cl9vcGNpb24oCiAgICAgICAgICAgICAgICAiICBPcGNpb24gWzFdOiAiLAogICAgICAgICAgICAgICAgeyIxIjogKCJTaSwg"
    "Z2VuZXJhciBlbCBpbmZvcm1lIiwgVHJ1ZSksCiAgICAgICAgICAgICAgICAgIjIiOiAoIk5vIiwgRmFsc2UpfSwKICAgICAg"
    "ICAgICAgICAgICIxIiwKICAgICAgICAgICAgKQoKICAgICAgICAgICAgaWYgaGFjZXJfaW5mb3JtZToKICAgICAgICAgICAg"
    "ICAgIG1hcmNhID0gZGF0ZXRpbWUubm93KCkuc3RyZnRpbWUoIiVZJW0lZF8lSCVNJVMiKQogICAgICAgICAgICAgICAgcnV0"
    "YV94bHN4ID0gb3MucGF0aC5qb2luKGNhcnBldGEsIGYiZGlzZW5vX2FjaTIxMV97bWFyY2F9Lnhsc3giKQogICAgICAgICAg"
    "ICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIGV4cG9ydGFyX2V4Y2VsKGQsIFIsIFMsIHJ1dGFfeGxzeCkKICAgICAg"
    "ICAgICAgICAgICAgICBwcmludChmIlxuICBbT0tdIEluZm9ybWUgZ3VhcmRhZG8gZW46XG4gICAgICAge3J1dGFfeGxzeH0i"
    "KQogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlcnJvcjoKICAgICAgICAgICAgICAgICAgICBwcmludChm"
    "IlxuICBbIV0gTm8gc2UgcHVkbyBnZW5lcmFyIGVsIGluZm9ybWU6IHtlcnJvcn0iKQogICAgICAgICAgICBlbHNlOgogICAg"
    "ICAgICAgICAgICAgcHJpbnQoIlxuICBJbmZvcm1lIG5vIGdlbmVyYWRvLiIpCgogICAgICAgICAgICBwYXVzYSgpCgogICAg"
    "ICAgIGVsaWYgb3BjaW9uID09ICIyIjoKICAgICAgICAgICAgcmV0dXJuCiAgICAgICAgZWxzZToKICAgICAgICAgICAgcHJp"
    "bnQoIlxuICBbIV0gT3BjaW9uIG5vIHZhbGlkYS4iKQogICAgICAgICAgICBwYXVzYSgpCgoKaWYgX19uYW1lX18gPT0gIl9f"
    "bWFpbl9fIjoKICAgIG1haW4oKQo="
)
