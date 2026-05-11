"""
Tests básicos de los endpoints requeridos por el Innovation Hub.
Verifica que /api/health, /api/metrics y /api/metrics/menu respondan correctamente.
"""
import pytest
from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)


def test_health():
    r = client.get("/api/health")
    assert r.status_code == 200
    body = r.json()
    assert body["status"] == "ok"
    assert "service" in body
    assert "timestamp" in body


def test_metrics():
    r = client.get("/api/metrics")
    assert r.status_code == 200
    body = r.json()
    assert "kpis" in body
    assert isinstance(body["kpis"], list)
    assert len(body["kpis"]) >= 1


def test_metrics_menu():
    r = client.get("/api/metrics/menu")
    assert r.status_code == 200
    body = r.json()
    assert "name" in body
    assert "links" in body
    assert isinstance(body["links"], list)


def test_root():
    r = client.get("/")
    assert r.status_code == 200
    body = r.json()
    assert "ui" in body
    assert "health" in body
