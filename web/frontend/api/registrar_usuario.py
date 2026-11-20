# -*- coding: utf-8 -*-
"""API serverless para VERCEL — POST /api/registrar_usuario (Equipo 6).

Solo el administrador (admin@aci211.com) puede registrar nuevos usuarios.
Usa la service role key de Supabase, que solo existe en el servidor
(NUNCA se expone al navegador).

Variables de entorno en Vercel:
  VITE_SUPABASE_URL          -> https://uiudnidegokdnnisjpho.supabase.co
  VITE_SUPABASE_ANON_KEY     -> clave publishable (verifica al llamante)
  SUPABASE_SERVICE_ROLE_KEY  -> service role key (secreta)
"""

import json
import os
import urllib.error
import urllib.request

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

app = FastAPI(title="Registrar usuario (admin)", version="1.0")

ADMIN_EMAIL = "admin@aci211.com"
MIN_PASS = 6


def _env(name):
    return (os.environ.get(name) or "").strip()


def _url():
    return _env("VITE_SUPABASE_URL") or _env("SUPABASE_URL")


def _anon():
    return _env("VITE_SUPABASE_ANON_KEY") or _env("SUPABASE_ANON_KEY")


def _service():
    return _env("SUPABASE_SERVICE_ROLE_KEY")


def _http(method, url, headers=None, body=None):
    data = None if body is None else json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, method=method)
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read().decode("utf-8", "replace")
            try:
                return r.status, json.loads(raw) if raw else {}
            except json.JSONDecodeError:
                return r.status, {"raw": raw}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            return e.code, {"raw": raw}
    except Exception as e:  # pragma: no cover
        return 0, {"error": str(e)}


@app.post("/api/registrar_usuario")
async def registrar(request: Request):
    url = _url()
    anon = _anon()
    service = _service()
    if not url or not anon or not service:
        return JSONResponse(
            {"ok": False, "detalle": "Faltan variables de entorno de Supabase en el servidor."},
            status_code=500,
        )

    # 1) Verificamos que el llamante sea el administrador
    auth = request.headers.get("authorization") or ""
    if not auth.lower().startswith("bearer "):
        return JSONResponse({"ok": False, "detalle": "No hay sesión (falta Authorization)."}, status_code=401)
    token = auth.split(" ", 1)[1].strip()

    status, user = _http("GET", f"{url}/auth/v1/user", headers={
        "Authorization": f"Bearer {token}",
        "apikey": anon,
    })
    if status != 200 or not isinstance(user, dict) or user.get("email") != ADMIN_EMAIL:
        return JSONResponse({"ok": False, "detalle": "Solo el administrador puede registrar usuarios."}, status_code=403)

    # 2) Leemos y validamos el cuerpo
    try:
        body = await request.json()
    except Exception:
        body = {}
    email = (body.get("email") or "").strip()
    password = body.get("password") or ""
    if not email or "@" not in email:
        return JSONResponse({"ok": False, "detalle": "Correo inválido."}, status_code=400)
    if len(password) < MIN_PASS:
        return JSONResponse({"ok": False, "detalle": f"La contraseña debe tener al menos {MIN_PASS} caracteres."}, status_code=400)

    # 3) Creamos el usuario con la service role key (confirmado automáticamente)
    status, created = _http("POST", f"{url}/auth/v1/admin/users", headers={
        "Authorization": f"Bearer {service}",
        "apikey": service,
    }, body={
        "email": email,
        "password": password,
        "email_confirm": True,
    })
    if status not in (200, 201):
        detalle = (
            created.get("msg")
            or created.get("message")
            or created.get("error_description")
            or created.get("error")
            or str(created)
        )
        return JSONResponse({"ok": False, "detalle": detalle}, status_code=status or 400)

    return {"ok": True, "email": created.get("email"), "id": created.get("id")}
