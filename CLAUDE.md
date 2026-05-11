# Agente Contable — Analizador de Libro Mayor

## ¿Qué hace este proyecto?

Servicio IA para **ComunidadFeliz SPA** (Tesorería) que analiza reportes de **Libro Mayor** (Excel `.xlsx`):

1. **Parsea** la cuenta contable del formato CF (encabezado en fila 0-1, datos desde fila detectada por "Cuenta"+"Fecha")
2. **Enriquece** filas sin RUT consultando el portal Clay en modo headless (Playwright)
3. **Normaliza** nombres de proveedores con IA (ej. "FACEBOOK" = "Facebook" = "fb ads")
4. **Aplica IVA 19%** solo a proveedores con RUT chileno; extranjeros/CLAY quedan con valor neto
5. **Genera informe Excel** con 2 hojas: Mayor limpio + Análisis (KPIs, tabla proveedores, distribución por período)
6. **Expone endpoints Hub** (`/api/health`, `/api/metrics`, `/api/metrics/menu`) para monitoreo

## Estructura

```
├── app/            FastAPI — endpoints del Innovation Hub
│   ├── main.py
│   └── routers/
│       ├── health.py    /api/health
│       └── metrics.py   /api/metrics + /api/metrics/menu
├── ui/
│   └── app.py      Streamlit — interfaz principal de usuario
├── scripts/
│   └── setup.py    Setup inicial del proyecto
├── tests/          pytest
├── Dockerfile
├── hub.config.json
└── .env            (NO commiteado — ver .env.example)
```

## Columnas del Libro Mayor CF

| Índice | Campo       |
|--------|-------------|
| 0      | Cuenta      |
| 1      | Fecha       |
| 2      | Nº Asiento  |
| 3      | Débito      |
| 4      | Crédito     |
| 5      | Acumulado   |
| 6      | Obligación  |
| 7      | RUT         |
| 8      | Contraparte |

## Regla IVA

- **Con RUT chileno** → `neto × 1.19` (afecto IVA, opera en Chile)
- **Sin RUT / RUT="CLAY"** → `neto × 1.0` (exento, proveedor extranjero)

## Variables de entorno requeridas

```
GROQ_API_KEY=   # Groq API para análisis IA (gratis en console.groq.com)
CLAY_EMAIL=     # Cuenta Clay (app.clay.cl) para enriquecimiento
CLAY_PASSWORD=  # Contraseña Clay
```

## Cómo iniciar (desarrollo)

```bash
# Setup inicial (solo primera vez)
python scripts/setup.py

# FastAPI (Hub monitoring)
uvicorn app.main:app --reload --port 8000

# Streamlit UI (herramienta principal)
python -m streamlit run ui/app.py --server.port 8501
```

## Tests

```bash
pytest tests/ -v
```

## Launcher Windows (producción)

Ver `Iniciar Agente Contable.vbs` en Google Drive → detecta si el puerto 8501 está ocupado antes de lanzar.

## Dependencias clave

- `groq` → llama-3.3-70b-versatile
- `playwright` → headless Chromium para Clay
- `openpyxl` → generación de Excel
- `streamlit` → UI principal
- `fastapi` → endpoints del Hub
