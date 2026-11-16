# Web React + API REST para el Diseño de Mezclas ACI 211

Arquitectura de 3 capas donde **el cerebro sigue siendo el script Python**:

```
┌─────────────────┐      ┌──────────────────┐      ┌──────────────────────────┐
│  React (Vite)   │ HTTP │  API FastAPI     │ stdin│ diseno_mezclas_aci211.py │
│  frontend/dist  │─────▶│  web/backend/    │─────▶│ (SIN MODIFICAR: escribe  │
│  formulario     │ JSON │  ejecuta el .py  │stdout│  por consola, genera el  │
│  + resultados   │◀─────│  y parsea salida │◀─────│  Excel y calcula TODO)   │
└─────────────────┘      └──────────────────┘      └──────────────────────────┘
```

El backend **no recalcula nada**: lanza `diseno_mezclas_aci211.py` como proceso
hijo, le envía las respuestas por su entrada estándar (igual que un usuario
tecleando, donde una línea vacía = ENTER = valor por defecto del ejemplo
"corregido"), lee su salida y la convierte a JSON. El Excel que descarga la web
es el mismo que genera el script.

## Cómo usar (presentación)

1. Doble clic en `iniciar_web.bat` → abre `http://127.0.0.1:8000`
2. El formulario ya trae los valores del ejemplo; pulsa **⚙️ Calcular dosificación**
3. Verás f'cr, a/c, dosis por m³, proporción 1 : 2.03 : 3.01, por saco y probetas
4. Botón para **descargar el Excel** (las 5 hojas + 3 gráficos del script)
5. Sección "Salida real del script" para demostrar que es el Python quien calcula

## Estructura

| Carpeta | Contenido |
|---|---|
| `backend/api.py` | API REST (FastAPI + uvicorn, ya instalados con pip) |
| `frontend/` | Proyecto React + Vite (código fuente) |
| `frontend/dist/` | Build de producción que sirve el backend |

## API

| Método | Ruta | Descripción |
|---|---|---|
| GET | `/api/health` | Estado del servicio |
| POST | `/api/disenar` | Ejecuta el script con los datos JSON y devuelve resultados |
| GET | `/api/informes/{nombre}` | Descarga el Excel generado |

Ejemplo rápido (PowerShell):

```powershell
Invoke-RestMethod -Uri http://127.0.0.1:8000/api/disenar -Method Post `
  -ContentType "application/json" -Body '{"fc": 250}'
```

Con solo `{"fc": 250}` el resto de campos van vacíos → el script usa sus
valores por defecto (ENTER en cada pregunta).

## Regenerar el frontend (solo si se edita)

```
cd frontend
npm install
npm run build
```

No es necesario para usar la web: el build ya está compilado en `dist/`.
