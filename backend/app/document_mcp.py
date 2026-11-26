"""Local stdio tools for a single, operator-selected pilot case."""

import argparse
import json
import os
import shutil
import subprocess
from tempfile import TemporaryDirectory
from functools import lru_cache
from pathlib import Path
from typing import Annotated, Literal
from uuid import UUID, uuid4
from zipfile import BadZipFile, ZipFile

from docx import Document
from mcp.server.fastmcp import FastMCP
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, range_boundaries
from pydantic import BaseModel, ConfigDict, Field
from pypdf import PdfReader


class FindingRow(BaseModel):
    model_config = ConfigDict(extra="forbid")
    location: str = Field(min_length=1, max_length=200)
    evidence: str = Field(min_length=1, max_length=400)
    observation: str = Field(min_length=1, max_length=500)
    recommendation: str = Field(min_length=1, max_length=500)
    status: Literal["sustentado", "requiere_verificacion"]


class ExcelFinding(FindingRow):
    sheet: str = Field(min_length=1, max_length=31)
    cell: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")


def export_marked_excel(source: Path, destination: Path, findings: list[ExcelFinding]) -> None:
    if os.name != "nt":
        raise ValueError("La exportación XLSX necesita Excel de escritorio en Windows")
    coordinates = {(f.sheet.casefold(), f.cell) for f in findings}
    if len(coordinates) != len(findings):
        raise ValueError("Consolide los hallazgos repetidos por celda")
    script = Path(__file__).resolve().parents[2] / "scripts" / "revisar-errores-excel.ps1"
    powershell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    with TemporaryDirectory() as temporary:
        folder = Path(temporary)
        staged = folder / "original.xlsx"
        output = folder / "revision.xlsx"
        plan = folder / "hallazgos.json"
        shutil.copyfile(source, staged)
        plan.write_text(json.dumps([
            {"Sheet": f.sheet, "Cell": f.cell, "Error": f.observation,
             "Evidence": f.evidence, "Recommendation": f.recommendation,
             "Status": "SUSTENTADO" if f.status == "sustentado" else "REQUIERE VERIFICACION"}
            for f in findings
        ]), encoding="utf-8")
        try:
            subprocess.run(
                [str(powershell), "-NoProfile", "-NonInteractive", "-File", str(script),
                 "-Libro", str(staged), "-Hallazgos", str(plan), "-Salida", str(output)],
                check=True, timeout=40, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        except (OSError, subprocess.SubprocessError):
            raise ValueError("No se exportó: verifica Excel y que todas las celdas sean visibles, no vacías y editables") from None
        with output.open("rb") as produced, destination.open("xb") as artifact:
            shutil.copyfileobj(produced, artifact)


def build_server(root: Path, case_id: UUID, *, write_reports: bool = False) -> FastMCP:
    root = root.resolve(strict=True)
    case = root / str(case_id)
    originals = case / "originals"
    if originals.resolve(strict=True) != originals or not originals.is_dir():
        raise ValueError("Directorio de originales no permitido")
    server = FastMCP("avaluos-informes" if write_reports else "avaluos-lectura")

    def document_path(document_id: UUID) -> Path:
        path = originals / str(document_id)
        try:
            allowed = path.resolve(strict=True) == path and path.is_file() and path.stat().st_size <= 64 * 1024**2
        except OSError:
            allowed = False
        if not allowed:
            raise ValueError("Documento no permitido")
        return path

    def kind(path: Path) -> str:
        with path.open("rb") as file:
            if file.read(5) == b"%PDF-":
                return "pdf"
            file.seek(0)
            with ZipFile(file) as zip_file:
                entries = zip_file.infolist()
                if len(entries) > 4096 or sum(e.file_size for e in entries) > 160 * 1024**2:
                    raise ValueError("Paquete de documento supera los límites")
                names = set(zip_file.namelist())
                if len(names) != len(entries) or any(n.endswith("vbaProject.bin") for n in names):
                    raise ValueError("Paquete no permitido")
                if "xl/workbook.xml" in names:
                    return "xlsx"
                if "word/document.xml" in names:
                    return "docx"
        raise ValueError("Formato no admitido")

    def require(document_id: UUID, expected: str) -> Path:
        path = document_path(document_id)
        if kind(path) != expected:
            raise ValueError("El tipo de documento no coincide")
        return path

    if write_reports:
        def report_directory() -> Path:
            reports = case / "reports"
            reports.mkdir(exist_ok=True)
            if reports.resolve(strict=True) != reports:
                raise ValueError("Directorio de informes no permitido")
            return reports

        @server.tool()
        def marcar_excel(document_id: UUID,
                         hallazgos: Annotated[list[ExcelFinding], Field(max_length=50)]) -> dict:
            """Anota hallazgos ya revisados en una copia XLSX. Nunca cambiar valores ni fórmulas."""
            source = require(document_id, "xlsx")
            report_id = uuid4()
            export_marked_excel(source, report_directory() / f"{report_id}.xlsx", hallazgos)
            return {"report_id": str(report_id), "marked_cells": len(hallazgos)}

        @server.tool()
        def crear_informe(hallazgos: Annotated[list[FindingRow], Field(max_length=50)],
                          resumen: Annotated[str, Field(max_length=1500)] = "",
                          cobertura: Annotated[list[str], Field(max_length=64)] | None = None) -> dict:
            """Crea un DOCX de revisión, separado del original. No certifica las conclusiones."""
            reports = report_directory()
            report_id = uuid4()
            doc = Document()
            doc.add_heading("Revisión técnica", 0)
            doc.add_paragraph("Borrador para revisión humana. No autoriza la entrega bancaria.")
            if resumen:
                doc.add_heading("Resumen del agente", level=1)
                doc.add_paragraph(resumen)
            if cobertura:
                doc.add_heading("Cobertura y límites", level=1)
                doc.add_paragraph("Los hallazgos requieren verificación independiente. El informe no acredita una revisión integral.")
                for item in cobertura:
                    doc.add_paragraph(item)
            for finding in hallazgos:
                doc.add_heading(finding.location, level=1)
                doc.add_paragraph(f"Estado: {finding.status}")
                doc.add_paragraph(f"Observación: {finding.observation}")
                doc.add_paragraph(f"Evidencia: {finding.evidence}")
                doc.add_paragraph(f"Recomendación: {finding.recommendation}")
            with (reports / f"{report_id}.docx").open("xb") as stream:
                doc.save(stream)
            return {"report_id": str(report_id), "findings": len(hallazgos)}

        return server

    @server.tool()
    def inventario() -> dict:
        """Lista identificadores y formatos, sin mostrar nombres de archivos originales."""
        result = []
        for path in originals.iterdir():
            if len(result) == 64:
                raise ValueError("El expediente supera 64 documentos")
            try:
                document_id = UUID(path.name)
                if path.name != str(document_id):
                    continue
            except ValueError:
                continue
            try:
                format_name = kind(document_path(document_id))
            except (OSError, ValueError, BadZipFile):
                format_name = "no_admitido"
            result.append({"document_id": str(document_id), "format": format_name})
        return {"documents": result}

    @lru_cache(maxsize=2)
    def workbook(document_id: UUID, cached: bool = False):
        with require(document_id, "xlsx").open("rb") as stream:
            return load_workbook(stream, data_only=cached, keep_links=False)

    @server.tool()
    def hojas_excel(document_id: UUID) -> dict:
        """Lista hojas de cálculo y su visibilidad; no devuelve valores de hojas ocultas."""
        return {"sheets": [{"sheet": sheet.title, "visible": sheet.sheet_state == "visible"}
                           for sheet in workbook(document_id).worksheets]}

    @server.tool()
    def leer_rango_excel(document_id: UUID, hoja: Annotated[str, Field(max_length=31)],
                         rango: Annotated[str, Field(max_length=24)]) -> dict:
        """Lee hasta 100 celdas visibles. Devuelve fórmula, formato y valor guardado; no recalcula."""
        left, top, right, bottom = range_boundaries(rango)
        if not (left and top and right and bottom and left <= right <= 16384 and top <= bottom <= 1048576
                and (right - left + 1) * (bottom - top + 1) <= 100):
            raise ValueError("Indique un rango válido de hasta 100 celdas")
        sheet = workbook(document_id)[hoja]
        if sheet.sheet_state != "visible":
            raise ValueError("La hoja está oculta")
        values = workbook(document_id, True)[hoja]
        hidden_columns = set()
        for dimension in sheet.column_dimensions.values():
            if dimension.hidden:
                start = dimension.min or column_index_from_string(dimension.index)
                hidden_columns.update(range(max(left, start), min(right, dimension.max or start) + 1))
        cells = []
        hidden = 0
        for row in sheet.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
            for cell in row:
                row_dimension = sheet.row_dimensions.get(cell.row)
                if (row_dimension and row_dimension.hidden) or cell.column in hidden_columns:
                    hidden += 1
                    continue
                value = values[cell.coordinate].value
                formula = cell.value if cell.data_type == "f" else None
                if formula is not None and not isinstance(formula, str):
                    formula = getattr(formula, "text", None)
                cells.append({"cell": cell.coordinate, "formula": clip(formula, 250),
                              "cached_value": clip(value, 250), "cached_type": values[cell.coordinate].data_type,
                              "number_format": clip(cell.number_format, 100)})
        return {"sheet": hoja, "cells": cells, "hidden_omitted": hidden,
                "text_limits": {"formula": 250, "cached_value": 250, "number_format": 100}}

    @server.tool()
    def leer_pagina_pdf(document_id: UUID, pagina: Annotated[int, Field(ge=1)],
                        offset: Annotated[int, Field(ge=0)] = 0) -> dict:
        """Extrae texto de una página física de PDF; no aplica OCR."""
        with require(document_id, "pdf").open("rb") as stream:
            reader = PdfReader(stream)
            if reader.is_encrypted or pagina > len(reader.pages):
                raise ValueError("PDF cifrado o página inexistente")
            text = reader.pages[pagina - 1].extract_text()
            return {"page": pagina, "page_count": len(reader.pages),
                    "searchable_text": bool(text.strip()), **window(text, offset)}

    @server.tool()
    def leer_bloque_docx(document_id: UUID, bloque: Annotated[int, Field(ge=1)],
                         fila: Annotated[int, Field(ge=1)] = 1,
                         offset: Annotated[int, Field(ge=0)] = 0) -> dict:
        """Lee un bloque del cuerpo DOCX, en orden: párrafo o fila de tabla; sin paginación ficticia."""
        with require(document_id, "docx").open("rb") as stream:
            blocks = list(Document(stream).iter_inner_content())
            if bloque > len(blocks):
                raise ValueError("Bloque inexistente")
            block = blocks[bloque - 1]
            result = {"block": bloque, "block_count": len(blocks)}
            if hasattr(block, "rows"):
                if fila > len(block.rows):
                    raise ValueError("Fila inexistente")
                text = " | ".join(cell.text for cell in block.rows[fila - 1].cells)
                result.update(row=fila, row_count=len(block.rows))
            else:
                text = block.text
            return {**result, **window(text, offset)}

    return server


def window(text: str, offset: int) -> dict:
    end = offset + 4000
    return {"text": text[offset:end], "next_offset": end if end < len(text) else None}


def clip(value, limit: int):
    if value is None or isinstance(value, (bool, int, float)):
        return value
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--case", required=True, type=UUID)
    parser.add_argument("--mode", choices=["read", "report"], default="read")
    args = parser.parse_args()
    build_server(args.root, args.case, write_reports=args.mode == "report").run(transport="stdio")


if __name__ == "__main__":
    main()
