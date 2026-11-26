from fastapi.testclient import TestClient

from app.main import app


def test_pilot_refuses_disabled_calls_cross_origin_and_document_bodies(monkeypatch):
    client = TestClient(app, base_url="http://127.0.0.1:5174")
    headers = {"Origin": "http://127.0.0.1:5174", "X-Avaluos-Pilot": "synthetic"}
    monkeypatch.delenv("AVALUOS_LOCAL_PILOT", raising=False)
    assert client.post("/api/pilot", headers=headers).status_code == 503
    monkeypatch.setenv("AVALUOS_LOCAL_PILOT", "1")
    assert client.post("/api/pilot", headers={**headers, "Origin": "https://example.org"}).status_code == 403
    assert client.post("/api/pilot", headers=headers, content=b"client document").status_code == 403
    assert client.get("/api/pilot").json() == {
        "enabled": True, "model": "openai-codex/gpt-6.1-terra", "effort": "medium",
    }
