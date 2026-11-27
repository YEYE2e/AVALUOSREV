from pathlib import Path

from docx import Document
from openpyxl import Workbook

from app.analyzer import (
    index_workbook,
    load_dual_workbook,
    read_docx_blocks,
    read_pdf_pages,
    read_sheet_rows,
    resolve_excel_target,
    search_docx,
    search_workbook,
    trace_cell_dependencies,
)


def test_full_workbook_analyzer_detects_errors_external_links_hidden_deps_and_formula_breaks(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Estudio"
    ws["A1"] = "Comparable"
    ws["B1"] = "Precio"
    ws["C1"] = "Area"
    ws["D1"] = "Unitario"

    # 4 rows: rows 2, 3, 5 have consistent formulas, row 4 has a hardcoded formula break!
    for r, price, area in ((2, 100000, 200), (3, 120000, 240), (4, 150000, 300), (5, 180000, 360)):
        ws[f"A{r}"] = f"Comp {r - 1}"
        ws[f"B{r}"] = price
        ws[f"C{r}"] = area
        ws[f"D{r}"] = 999 if r == 4 else f"=B{r}/C{r}"

    # Far cell way past A1:L32 with external link
    ws["Z90"] = "='[OtroLibro.xlsx]Base'!A1"
    ws["Y90"] = "Valor externo"

    # Hidden sheet feeding a visible cell
    hidden = wb.create_sheet("CalculoOculto")
    hidden.sheet_state = "hidden"
    hidden["B2"] = "#REF!"
    ws["E10"] = "=CalculoOculto!B2"
    ws["D10"] = "Total oculto"

    # Merged cell
    ws["G20"] = "Resumen final"
    ws.merge_cells("G20:I20")

    path = tmp_path / "full.xlsx"
    wb.save(path)

    loaded = load_dual_workbook(path)
    info = index_workbook(path, loaded=loaded)
    assert info["visible_sheets"] == 1
    assert info["hidden_sheets"] == 1
    alerts = info["deterministic_alerts"]
    assert any(x["cell"] == "Z90" for x in alerts["external_links"])
    assert any(x["cell"] == "D4" and x["hardcoded_value"] == 999 for x in alerts["formula_breaks"])
    assert any(x["cell"] == "E10" for x in alerts["hidden_dependencies"])

    # Full-workbook search finds far cell Z90/Y90
    res = search_workbook(loaded, "Valor externo")
    assert res["count"] == 1 and res["matches"][0]["cell"] == "Y90"

    # Dependency tracer identifies hidden precedent on E10
    trace = trace_cell_dependencies(loaded, "Estudio", "E10")
    assert any(p["sheet"] == "CalculoOculto" and p["hidden"] is True for p in trace["precedents"])

    # Merged cell H20 resolves deterministically to anchor G20
    rs, rc, err = resolve_excel_target(path, "Estudio", "H20")
    assert err is None and rs == "Estudio" and rc == "G20"

    # Hidden sheet rejected
    rs2, rc2, err2 = resolve_excel_target(path, "CalculoOculto", "B2")
    assert rs2 is None and err2 is not None


def test_docx_search_covers_paragraphs_tables_and_headers(tmp_path: Path):
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Encabezado Catastro 2026"
    doc.add_paragraph("Escritura pública del predio urbano.")
    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "Clave catastral"
    tbl.cell(0, 1).text = "01-02-03-045"
    tbl.cell(1, 0).text = "Superficie"
    tbl.cell(1, 1).text = "425.50 m2"
    path = tmp_path / "respaldo.docx"
    doc.save(path)

    r1 = search_docx(path, "01-02-03-045")
    assert r1["count"] == 1 and r1["matches"][0]["type"] == "table" and r1["matches"][0]["row"] == 1
    r2 = search_docx(path, "Encabezado Catastro")
    assert r2["count"] == 1 and "header" in r2["matches"][0]["type"]


def test_formula_inconsistencies_and_bulk_readers(tmp_path: Path):
    from pypdf import PdfWriter

    wb = Workbook()
    ws = wb.active
    ws.title = "Mercado"
    ws["A1"] = "Ref"
    ws["B1"] = "Precio"
    ws["C1"] = "Area"
    ws["D1"] = "Otro"
    ws["E1"] = "Unitario"
    for r in range(2, 7):
        ws[f"A{r}"] = f"Ref {r - 1}"
        ws[f"B{r}"] = 100000 + r * 1000
        ws[f"C{r}"] = 200
        ws[f"D{r}"] = 100
        # Row 4 breaks the column formula pattern by dividing by D instead of C
        ws[f"E{r}"] = f"=B{r}/D{r}" if r == 4 else f"=B{r}/C{r}"
    ws.row_dimensions[3].hidden = True

    xlsx = tmp_path / "mercado.xlsx"
    wb.save(xlsx)
    loaded = load_dual_workbook(xlsx)
    info = index_workbook(xlsx, loaded=loaded)
    inconsistencies = info["deterministic_alerts"]["formula_inconsistencies"]
    assert any(x["cell"] == "E4" and x["formula"] == "=B4/D4" for x in inconsistencies)

    # Bulk row reader skips hidden row 3 and supports pagination
    page1 = read_sheet_rows(loaded, "Mercado", start_row=1, max_rows=2, max_cells=600)
    assert [r["row"] for r in page1["rows"]] == [1, 2]
    assert page1["next_row"] == 4
    page2 = read_sheet_rows(loaded, "Mercado", start_row=page1["next_row"], max_rows=10, max_cells=600)
    assert [r["row"] for r in page2["rows"]] == [4, 5, 6]
    assert page2["next_row"] is None

    # Bulk DOCX reader reads header + paragraphs + full table in one call
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Certificado Registral"
    doc.add_paragraph("Propietario: Juan Pérez")
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "Predio"
    t.cell(0, 1).text = "45678"
    t.cell(1, 0).text = "Gravamen"
    t.cell(1, 1).text = "Hipoteca abierta"
    docx_path = tmp_path / "cert.docx"
    doc.save(docx_path)
    docx_bulk = read_docx_blocks(docx_path, start_block=1, max_blocks=50)
    assert any("Certificado Registral" in h["text"] for h in docx_bulk["headers_footers"])
    assert len(docx_bulk["blocks"]) == 2
    assert docx_bulk["blocks"][1]["type"] == "table"
    assert docx_bulk["blocks"][1]["rows"][1]["text"] == "Gravamen | Hipoteca abierta"
    assert docx_bulk["next_block"] is None

    # Bulk PDF page reader covers multiple pages cleanly
    pdf_path = tmp_path / "multi.pdf"
    writer = PdfWriter()
    for _ in range(3):
        writer.add_blank_page(width=100, height=100)
    writer.write(pdf_path)
    pdf_bulk = read_pdf_pages(pdf_path, start_page=1, end_page=3)
    assert pdf_bulk["page_count"] == 3
    assert len(pdf_bulk["pages"]) == 3
    assert pdf_bulk["next_page"] is None
