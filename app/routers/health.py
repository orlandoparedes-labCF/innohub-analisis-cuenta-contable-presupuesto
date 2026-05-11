"""Health endpoint — requerido por el Innovation Hub para monitoreo."""
from fastapi import APIRouter
from datetime import datetime, timezone

router = APIRouter()


@router.get("/health")
async def health():
    return {
        "status": "ok",
        "service": "analisis-cuenta-contable-presupuesto",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }
