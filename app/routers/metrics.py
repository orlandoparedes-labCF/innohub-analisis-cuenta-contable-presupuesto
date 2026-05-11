"""Metrics endpoints — requerido por el Innovation Hub para dashboard y menú."""
from fastapi import APIRouter
from datetime import datetime, timezone
import json
from pathlib import Path

router = APIRouter()

# Ruta al archivo de estado persistente (se crea en runtime)
STATE_FILE = Path(__file__).parent.parent.parent / "data" / "metrics_state.json"


def _load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text())
        except Exception:
            pass
    return {"analyses_total": 0, "last_analysis": None, "last_account": None}


@router.get("/metrics")
async def metrics():
    """Métricas del servicio para el dashboard del Hub."""
    state = _load_state()
    return {
        "service": "analisis-cuenta-contable-presupuesto",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "kpis": [
            {
                "key": "analyses_total",
                "label": "Análisis realizados",
                "value": state.get("analyses_total", 0),
                "unit": "análisis",
            },
            {
                "key": "last_analysis",
                "label": "Último análisis",
                "value": state.get("last_analysis") or "—",
                "unit": "fecha",
            },
            {
                "key": "last_account",
                "label": "Última cuenta analizada",
                "value": state.get("last_account") or "—",
                "unit": "cuenta",
            },
        ],
        "status": "active",
    }


@router.get("/metrics/menu")
async def metrics_menu():
    """Estructura de menú para el Innovation Hub."""
    return {
        "service": "analisis-cuenta-contable-presupuesto",
        "name": "Analizador Libro Mayor",
        "icon": "📊",
        "description": "Analiza cuentas contables del Libro Mayor con IA: proveedores, IVA, distribución por período.",
        "links": [
            {
                "label": "Abrir herramienta",
                "url": "http://localhost:8501",
                "type": "primary",
                "icon": "🚀",
            },
            {
                "label": "Health check",
                "url": "/api/health",
                "type": "internal",
            },
        ],
        "tags": ["contabilidad", "ia", "presupuesto", "libro-mayor", "tesoreria"],
    }
