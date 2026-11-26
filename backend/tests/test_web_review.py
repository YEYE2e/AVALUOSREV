import json
from io import BytesIO
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient

PdfWriter = pytest.importorskip("pypdf").PdfWriter
Workbook = pytest.importorskip("openpyxl").Workbook

from app import web_review
from app.main import app


def test_review_stages_upload_without_treating_blank_pdf_as_reviewed(tmp_path, monkeypatch):
    monkeypatch.setenv("AVALUOS_STORAGE_ROOT", str(tmp_path))
    client = TestClient(app, base_url="http://127.0.0.1:5174", client=("127.0.0.1", 50000))
    buffer = BytesIO()
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    writer.write(buffer)
    payload = buffer.getvalue()
    files = {"principal": ("anexo.pdf", payload, "application/pdf")}
    headers = {"Origin": "http://127.0.0.1:5174", "X-Avaluos-Review": "local"}
    assert client.post("/api/reviews/pdf", files=files).status_code == 403
    assert not list(tmp_path.iterdir())
    result = client.post("/api/reviews/pdf", headers=headers, files=files)
    assert result.status_code == 422
    assert "requiere OCR" in result.json()["detail"]
    originals = list(tmp_path.glob("*/originals/*"))
    assert len(originals) == 1 and originals[0].read_bytes() == payload
    assert not list(tmp_path.glob("*/reports/*"))


def test_run_agent_strips_markdown_fences_and_validates_schema(tmp_path, monkeypatch):
    payload = {
        "summary": "Resumen de prueba.",
        "findings": [
            {
                "location": "Mercado!C2",
                "evidence": "Mercado!B2 = 100",
                "observation": "Revisar cálculo.",
                "recommendation": "Validar fórmula.",
                "sheet": "Mercado",
                "cell": "C2",
            }
        ],
    }

    def fake_run(*args, **kwargs):
        return SimpleNamespace(returncode=0, stdout=f"```json\n{json.dumps(payload)}\n```")

    monkeypatch.setattr(web_review.subprocess, "run", fake_run)
    review = web_review.run_agent(tmp_path, "excel", {"documents": [], "coverage": []})
    assert review.summary == "Resumen de prueba."
    assert len(review.findings) == 1
    assert review.findings[0].sheet == "Mercado"
    assert review.findings[0].cell == "C2"


def test_excel_review_generates_word_and_marked_excel_and_keeps_original(tmp_path, monkeypatch):
    monkeypatch.setenv("AVALUOS_STORAGE_ROOT", str(tmp_path))
    book = Workbook()
    sheet = book.active
    sheet.title = "Mercado"
    sheet["A1"] = "Precio"
    sheet["B1"] = 1000
    sheet["C1"] = "=B1/2"
    excel_buf = BytesIO()
    book.save(excel_buf)
    excel_bytes = excel_buf.getvalue()

    def fake_run(*args, **kwargs):
        payload = {
            "summary": "Se detectó una observación en Mercado!C1.",
            "findings": [
                {
                    "location": "Mercado!C1",
                    "evidence": "Mercado!B1 = 1000",
                    "observation": "Verificar divisor",
                    "recommendation": "Confirmar área",
                    "sheet": "Mercado",
                    "cell": "C1",
                }
            ],
        }
        return SimpleNamespace(returncode=0, stdout=json.dumps(payload))

    monkeypatch.setattr(web_review.subprocess, "run", fake_run)
    client = TestClient(app, base_url="http://127.0.0.1:5174", client=("127.0.0.1", 50000))
    headers = {"Origin": "http://127.0.0.1:5174", "X-Avaluos-Review": "local"}
    response = client.post(
        "/api/reviews/excel",
        headers=headers,
        files={"principal": ("avaluo.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["findings"] == 1
    assert body["downloads"]["word"].endswith("/docx")
    word_resp = client.get(body["downloads"]["word"])
    assert word_resp.status_code == 200
    assert word_resp.content.startswith(b"PK")
    if "excel" in body["downloads"]:
        xlsx_resp = client.get(body["downloads"]["excel"])
        assert xlsx_resp.status_code == 200
        assert xlsx_resp.content.startswith(b"PK")
    originals = list(tmp_path.glob("*/originals/*"))
    assert len(originals) == 1 and originals[0].read_bytes() == excel_bytes
