import asyncio
import json
import sys
from pathlib import Path
from uuid import uuid4
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest

pytest.importorskip("mcp", reason="Instalar el extra documents para probar MCP")

from docx import Document
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from pypdf import PdfWriter

from app.workspace import Workspace


def test_case_scoped_stdio_readers_and_separate_writer(tmp_path):
    case, other = uuid4(), uuid4()
    xlsx, docx, pdf, foreign = (uuid4() for _ in range(4))
    workspace = Workspace(tmp_path)

    book = Workbook()
    sheet = book.active
    sheet.title = "Mercado"
    sheet.append([120000, 300, "=A1/B1"])
    sheet["C1"].comment = Comment("Nota original del revisor.", "Revisor")
    sheet.append(["omitido", "omitido", "omitido"])
    sheet.row_dimensions[2].hidden = True
    sheet.column_dimensions["A"].hidden = True
    # Add a far cell beyond A1:L32 to prove unlimited-size search & dependency tracing
    sheet["U126"] = "=C1*2"
    sheet["T126"] = "Valor proyectado"
    book.create_sheet("Aux").sheet_state = "hidden"
    book["Aux"]["A1"] = "#REF!"
    excel_file = tmp_path / "fixture.xlsx"
    book.save(excel_file)
    original = workspace.save_original(case, xlsx, excel_file.read_bytes())
    workspace.save_original(other, foreign, excel_file.read_bytes())

    word_file = tmp_path / "fixture.docx"
    word = Document()
    word.add_paragraph("Catastro ficticio")
    word.add_table(rows=1, cols=1).cell(0, 0).text = "300 m2"
    word.save(word_file)
    workspace.save_original(case, docx, word_file.read_bytes())

    pdf_file = tmp_path / "fixture.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    writer.write(pdf_file)
    workspace.save_original(case, pdf, pdf_file.read_bytes())

    def parameters(mode):
        return StdioServerParameters(
            command=sys.executable,
            cwd=Path(__file__).resolve().parents[1],
            args=["-m", "app.document_mcp", "--root", str(tmp_path), "--case", str(case), "--mode", mode],
        )

    async def exercise():
        async with stdio_client(parameters("read")) as (read, write):
            async with ClientSession(read, write) as client:
                await client.initialize()
                tools = {tool.name for tool in (await client.list_tools()).tools}
                assert "crear_informe" not in tools
                assert {"inventario", "resumen_expediente", "buscar_excel", "rastrear_celda_excel", "buscar_pdf", "buscar_docx", "ver_pagina_pdf", "ver_imagen"} <= tools

                async def call(tool, **args):
                    result = await client.call_tool(tool, args)
                    assert not result.isError, result.content
                    return json.loads(result.content[0].text)

                inventory = await call("inventario")
                assert {d["document_id"] for d in inventory["documents"]} == set(map(str, [xlsx, docx, pdf]))
                denied = await client.call_tool("hojas_excel", {"document_id": str(foreign)})
                assert denied.isError
                hidden = await client.call_tool("leer_rango_excel", {"document_id": str(xlsx), "hoja": "Aux", "rango": "A1"})
                assert hidden.isError

                rng = await call("leer_rango_excel", document_id=str(xlsx), hoja="Mercado", rango="A1:C2")
                assert [c["cell"] for c in rng["cells"]] == ["B1", "C1"]
                assert rng["cells"][1]["formula"] == "=A1/B1" and rng["cells"][1]["cached_value"] is None

                # Search far beyond A1:L32
                found = await call("buscar_excel", document_id=str(xlsx), consulta="Valor proyectado")
                assert found["count"] == 1 and found["matches"][0]["cell"] == "T126"

                # Trace precedents (including hidden A1!) and visible dependents (U126!)
                trace = await call("rastrear_celda_excel", document_id=str(xlsx), hoja="Mercado", celda="C1")
                assert any(p["cell"] == "A1" and p["hidden"] is True for p in trace["precedents"])
                assert any(d["cell"] == "U126" for d in trace["visible_dependents"])

                atlas = await call("resumen_expediente")
                assert len(atlas["documents"]) == 3

                table = await call("leer_bloque_docx", document_id=str(docx), bloque=2, fila=1)
                assert table["text"] == "300 m2"
                docx_search = await call("buscar_docx", document_id=str(docx), consulta="300 m2")
                assert docx_search["count"] == 1

                page = await call("leer_pagina_pdf", document_id=str(pdf), pagina=1)
                assert not page["searchable_text"]

        async with stdio_client(parameters("report")) as (read, write):
            async with ClientSession(read, write) as client:
                await client.initialize()
                assert {tool.name for tool in (await client.list_tools()).tools} == {"crear_informe", "marcar_excel"}
                finding = {
                    "sheet": "Mercado",
                    "cell": "C1",
                    "location": "Mercado!C1",
                    "evidence": "Mercado!B1",
                    "observation": "Revisar área",
                    "recommendation": "Cotejar el área",
                    "status": "sustentado",
                    "severity": "critico",
                    "review_type": "Cálculo / fórmula",
                    "value_impact": "demostrado",
                }
                denied = await client.call_tool("marcar_excel", {"document_id": str(xlsx), "hallazgos": [{**finding, "cell": "A1"}]})
                assert denied.isError
                marked = await client.call_tool("marcar_excel", {"document_id": str(xlsx), "hallazgos": [finding]})
                assert not marked.isError, marked.content
                marked_id = json.loads(marked.content[0].text)["report_id"]
                result = await client.call_tool(
                    "crear_informe",
                    {
                        "hallazgos": [{k: v for k, v in finding.items() if k not in ("sheet", "cell")}],
                        "resumen": "Resumen técnico",
                        "verified_independently": True,
                    },
                )
                assert not result.isError, result.content
                return json.loads(result.content[0].text)["report_id"], marked_id

    report_id, marked_id = asyncio.run(exercise())
    marked = load_workbook(tmp_path / str(case) / "reports" / f"{marked_id}.xlsx")
    cell = marked["Mercado"]["C1"]
    assert cell.value == "=A1/B1"
    assert cell.fill.fgColor.rgb.endswith("FF0000")
    with ZipFile(tmp_path / str(case) / "reports" / f"{marked_id}.xlsx") as archive:
        comments = [
            node.findtext("{*}text")
            for part in archive.namelist()
            if part.startswith("xl/threadedComments/") and part.endswith(".xml")
            for node in ElementTree.fromstring(archive.read(part)).findall("{*}ThreadedComment")
            if node.get("ref") == "C1"
        ]
        if not comments:
            comments = [cell.comment.text] if cell.comment else []
            assert any(
                node.findtext("{*}Row") == "0"
                and node.findtext("{*}Column") == "2"
                and node.find("{*}Visible") is not None
                for part in archive.namelist()
                if part.endswith(".vml")
                for node in ElementTree.fromstring(archive.read(part)).iterfind(".//{*}ClientData")
            )
    assert any(
        message and all(text in message for text in ("Nota original del revisor.", "Revisar área", "Mercado!B1", "Cotejar el área"))
        for message in comments
    )
    assert marked.active.title == "REVISIÓN TÉCNICA"
    review_sheet = marked["REVISIÓN TÉCNICA"]
    assert review_sheet["A1"].value == "REVISIÓN TÉCNICA DE AVALÚO — OBSERVACIONES PARA VALIDACIÓN DEL PERITO"
    assert "(Críticas: 1 |" in review_sheet["A2"].value
    assert [review_sheet.cell(5, col).value for col in range(1, 17)] == [
        "N.º",
        "Importancia",
        "Tipo de revisión",
        "Hoja",
        "Celda",
        "Dato revisado",
        "Hallazgo",
        "Explicación técnica",
        "Valor actual",
        "Valor referencia",
        "Diferencia",
        "Corrección recomendada",
        "Incidencia en valor",
        "Estado",
        "Evidencias",
        "Ubicación",
    ]
    assert review_sheet["P6"].hyperlink.location == "'Mercado'!C1"
    assert review_sheet["G6"].value == "Revisar área"
    report = Document(tmp_path / str(case) / "reports" / f"{report_id}.docx")
    assert any("Revisar área" in p.text for p in report.paragraphs)
    assert original.read_bytes() == excel_file.read_bytes()
