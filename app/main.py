"""
Innovation Hub — Analizador de Cuenta Contable y Presupuesto
FastAPI app principal. Expone los endpoints de monitoreo que requiere el Hub.
La interfaz de usuario corre como Streamlit en puerto 8501 (ver ui/app.py).
"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routers import health, metrics

app = FastAPI(
    title="Analizador Libro Mayor — ComunidadFeliz",
    description=(
        "Agente IA que analiza reportes de Libro Mayor: identifica proveedores, "
        "aplica IVA 19% a proveedores nacionales, enriquece via Clay, "
        "y genera informes Excel multi-período."
    ),
    version="1.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    openapi_url="/api/openapi.json",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(health.router, prefix="/api", tags=["monitoring"])
app.include_router(metrics.router, prefix="/api", tags=["monitoring"])


@app.get("/")
async def root():
    return {
        "service": "analisis-cuenta-contable-presupuesto",
        "ui": "http://localhost:8501",
        "docs": "/api/docs",
        "health": "/api/health",
    }
