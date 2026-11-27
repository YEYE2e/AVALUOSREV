import json
import time
from io import BytesIO

import pytest
from fastapi.testclient import TestClient

PdfWriter = pytest.importorskip("pypdf").PdfWriter
Workbook = pytest.importorskip("openpyxl").Workbook

from app import analyzer, orchestrator
from app.main import app


def test_review_stages_upload_and_rejects_unreadable_pdf_when_visual_unavailable(tmp_path, monkeypatch):
    monkeypatch.setenv("AVALUOS_STORAGE_ROOT", str(tmp_path))
    monkeypatch.setattr(analyzer, "pdftoppm_available", lambda: False)
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


def test_extract_json_object_handles_fences_and_surrounding_prose():
    raw = 'Aquí está el resultado:\n```json\n{"summary": "Ok", "findings": []}\n```\nFin.'
    extracted = orchestrator._extract_json_object(raw)
    parsed = json.loads(extracted)
    assert parsed["summary"] == "Ok"


def test_multi_stage_excel_pipeline_with_independent_verifier_and_history(tmp_path, monkeypatch):
    monkeypatch.setenv("AVALUOS_STORAGE_ROOT", str(tmp_path))
    monkeypatch.setenv("AVALUOS_ENABLE_VERIFIER", "1")

    book = Workbook()
    ws = book.active
    ws.title = "Valoracion"
    ws["A1"] = "Terreno"
    ws["B1"] = 1000
    # Cell way past A1:L32 + merged range to verify unlimited coverage & anchor resolution
    ws["U126"] = "=B1/2"
    ws.merge_cells("U126:V126")
    excel_buf = BytesIO()
    book.save(excel_buf)
    excel_bytes = excel_buf.getvalue()

    calls = []

    def fake_run_mcp_agent(*, verifier=False, **kwargs):
        calls.append(verifier)
        if not verifier:
            return json.dumps({
                "category_detected": "Urbano residencial",
                "summary": "Se hallaron dos candidatos en Valoracion.",
                "findings": [
                    {
                        "id": "H-01",
                        "severity": "critico",
                        "review_type": "Cálculo / fórmula",
                        "location": "Valoracion!V126",
                        "sheet": "Valoracion",
                        "cell": "V126",
                        "field_reviewed": "Valor unitario",
                        "observation": "Divisor incorrecto en fórmula",
                        "technical_explanation": "Divide por 2 en lugar del área real.",
                        "current_value": "=B1/2",
                        "expected_value": "=B1/4",
                        "difference": "250 USD",
                        "evidence": "Valoracion!U126 contiene =B1/2",
                        "recommendation": "Ajustar divisor al área verificada.",
                        "value_impact": "demostrado",
                        "status": "requiere_verificacion",
                    },
                    {
                        "id": "H-02",
                        "severity": "menor",
                        "review_type": "Coherencia interna",
                        "location": "Valoracion!A1",
                        "sheet": "Valoracion",
                        "cell": "A1",
                        "observation": "Falso positivo de etiqueta",
                        "evidence": "Valoracion!A1 = Terreno",
                        "recommendation": "Sin acción",
                        "value_impact": "ninguno_demostrado",
                        "status": "requiere_verificacion",
                    },
                ],
                "limitations": ["Sin escrituras adjuntas."],
            })
        return json.dumps({
            "verification_summary": "Se confirmó H-01 y se descartó H-02.",
            "decisions": [
                {
                    "finding_id": "H-01",
                    "decision": "sustentado",
                    "verification_note": "Comprobado directamente en Valoracion!U126.",
                },
                {
                    "finding_id": "H-02",
                    "decision": "descartado",
                    "verification_note": "Etiqueta estándar sin error.",
                },
            ],
            "additional_findings": [],
            "additional_limitations": [],
        })

    monkeypatch.setattr(orchestrator, "run_mcp_agent", fake_run_mcp_agent)
    client = TestClient(app, base_url="http://127.0.0.1:5174", client=("127.0.0.1", 50000))
    headers = {"Origin": "http://127.0.0.1:5174", "X-Avaluos-Review": "local"}

    # Test async job flow
    start = client.post(
        "/api/reviews/jobs/excel",
        headers=headers,
        files={"principal": ("avaluo-grande.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert start.status_code == 202
    job_id = start.json()["job_id"]
    case_id = start.json()["case_id"]

    snap = None
    for _ in range(100):
        resp = client.get(f"/api/reviews/jobs/{job_id}")
        assert resp.status_code == 200
        snap = resp.json()
        if snap["status"] in ("completed", "failed", "cancelled"):
            break
        time.sleep(0.1)

    assert snap and snap["status"] == "completed", snap
    assert calls == [False, True]  # Both primary auditor AND independent verifier ran!
    body = snap["result"]
    assert body["verified_independently"] is True
    assert body["category_detected"] == "Urbano residencial"
    assert body["findings"] == 1
    assert len(body["discarded"]) == 1
    # Merged cell V126 must have been resolved deterministically to anchor U126!
    assert body["items"][0]["cell"] == "U126"
    assert body["items"][0]["status"] == "sustentado"
    # Deterministic safety guards: 0 supports -> no_verificable_sin_respaldos; critical sustentado -> requiere_correcciones_tecnicas
    assert body["documentary_verdict"] == "no_verificable_sin_respaldos"
    assert body["conclusion_category"] in ("requiere_correcciones_tecnicas", "requiere_documentacion_o_verificacion")
    assert body["conclusion_label"]

    # Downloads work
    word_resp = client.get(body["downloads"]["word"])
    assert word_resp.status_code == 200 and word_resp.content.startswith(b"PK")
    if "excel" in body["downloads"]:
        xlsx_resp = client.get(body["downloads"]["excel"])
        assert xlsx_resp.status_code == 200 and xlsx_resp.content.startswith(b"PK")

    # History and persisted case result work
    hist = client.get("/api/reviews/history")
    assert hist.status_code == 200
    assert any(h["case_id"] == case_id and h["principal_label"] == "avaluo-grande.xlsx" for h in hist.json()["history"])
    saved = client.get(f"/api/reviews/{case_id}/result")
    assert saved.status_code == 200 and saved.json()["findings"] == 1

    # Original file remained strictly untouched
    originals = list(tmp_path.glob("*/originals/*"))
    assert len(originals) == 1 and originals[0].read_bytes() == excel_bytes


def test_deterministic_verdict_and_conclusion_safety_guards():
    from app.models import AgentReviewDraft, StructuredFinding

    crit_doc = StructuredFinding(
        id="H-01",
        severity="critico",
        review_type="Documental",
        location="Datos!B4",
        sheet="Datos",
        cell="B4",
        observation="Área difiere de escritura.",
        evidence="Datos!B4 vs Escritura pág. 3",
        recommendation="Ajustar área.",
        value_impact="demostrado",
        status="sustentado",
    )
    optimistic_draft = AgentReviewDraft(
        summary="Borrador optimista",
        documentary_verdict="concuerda",
        conclusion_category="conforme_para_revision_del_perito",
        findings=[crit_doc],
    )

    # With supports and a documentary discrepancy: cannot say concuerda or conforme
    doc_v, concl = orchestrator._resolve_verdicts(
        mode="excel",
        supports_count=2,
        draft=optimistic_draft,
        active_findings=[crit_doc],
        verified_ok=True,
    )
    assert doc_v == "existen_diferencias"
    assert concl == "requiere_correcciones_tecnicas"

    # Without supports in Excel: must say no_verificable_sin_respaldos
    doc_v2, _ = orchestrator._resolve_verdicts(
        mode="excel",
        supports_count=0,
        draft=optimistic_draft,
        active_findings=[crit_doc],
        verified_ok=True,
    )
    assert doc_v2 == "no_verificable_sin_respaldos"

    # In PDF mode: documentary verdict must always be no_aplica
    doc_v3, _ = orchestrator._resolve_verdicts(
        mode="pdf",
        supports_count=0,
        draft=optimistic_draft,
        active_findings=[],
        verified_ok=False,
    )
    assert doc_v3 == "no_aplica"
