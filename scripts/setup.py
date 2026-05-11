#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Setup del proyecto Innovation Hub - Analizador Libro Mayor.
Valida configuracion, crea carpetas necesarias, y verifica dependencias.
"""
import json
import os
import sys
import io

# Forzar UTF-8 en stdout para que los emojis no rompan en consolas cp1252
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pathlib import Path

ROOT = Path(__file__).parent.parent


def check_env():
    """Verifica que las variables de entorno requeridas estén configuradas."""
    env_file = ROOT / ".env"
    if not env_file.exists():
        example = ROOT / ".env.example"
        if example.exists():
            import shutil
            shutil.copy(example, env_file)
            print("✅ .env creado desde .env.example — recuerda completar los valores.")
        else:
            env_file.write_text(
                "GROQ_API_KEY=YOUR_GROQ_API_KEY\n"
                "CLAY_EMAIL=YOUR_CLAY_EMAIL\n"
                "CLAY_PASSWORD=YOUR_CLAY_PASSWORD\n"
            )
            print("✅ .env creado con plantilla vacía — completa los valores.")
    else:
        print("✅ .env ya existe.")


def create_data_dir():
    """Crea directorio de datos para estado persistente."""
    data_dir = ROOT / "data"
    data_dir.mkdir(exist_ok=True)
    state_file = data_dir / "metrics_state.json"
    if not state_file.exists():
        state_file.write_text(json.dumps({
            "analyses_total": 0,
            "last_analysis": None,
            "last_account": None,
        }, indent=2))
    print("✅ Directorio data/ listo.")


def validate_hub_config():
    """Valida que hub.config.json no tenga valores pendientes."""
    cfg_file = ROOT / "hub.config.json"
    if not cfg_file.exists():
        print("❌ hub.config.json no encontrado.")
        return False
    cfg = json.loads(cfg_file.read_text())
    text = json.dumps(cfg)
    if "CAMBIAR" in text:
        print("❌ hub.config.json tiene valores 'CAMBIAR' pendientes.")
        return False
    print(f"✅ hub.config.json OK — proyecto: {cfg['project']['name']}")
    return True


def check_python():
    """Verifica versión de Python."""
    if sys.version_info < (3, 10):
        print(f"⚠️  Python {sys.version} — se recomienda 3.10+")
    else:
        print(f"✅ Python {sys.version.split()[0]}")


def main():
    print("=" * 55)
    print("  Innovation Hub — Setup: Analizador Libro Mayor")
    print("=" * 55)
    check_python()
    check_env()
    create_data_dir()
    ok = validate_hub_config()
    print("=" * 55)
    if ok:
        print("✅ Setup completo. Para iniciar:")
        print("   API:  uvicorn app.main:app --reload --port 8000")
        print("   UI:   python -m streamlit run ui/app.py --server.port 8501")
    else:
        print("⚠️  Setup incompleto. Revisa los errores arriba.")
    print("=" * 55)


if __name__ == "__main__":
    main()
