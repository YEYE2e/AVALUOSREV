"""Deterministic full-document indexer, search engine, Excel dependency tracer, and visual renderer."""

from __future__ import annotations

import json
import re
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from uuid import UUID
from zipfile import BadZipFile, ZipFile

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_to_tuple,
    get_column_letter,
    range_boundaries,
)
from pypdf import PdfReader

EXCEL_ERRORS = frozenset({"#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"})
IMAGE_FORMATS = {"png": "image/png", "jpeg": "image/jpeg", "webp": "image/webp"}
MAX_ZIP_ENTRIES = 8192
MAX_ZIP_UNCOMPRESSED = 256 * 1024**2
MAX_FILE_BYTES = 128 * 1024**2

# Matches Excel cell/range references like 'Sheet 1'!$A$1:$B$10, Hoja!C5, or $D$12
_STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')
_REF_TOKEN_RE = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.\u00C0-\u024F]+))!)?"
    r"\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})"
    r"(?::\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6}))?"
)
_EXTERNAL_LINK_RE = re.compile(r"\[[^\]]+\]")
_CELL_COORD_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")


def clip(value: Any, limit: int = 250) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def sanitize_label(filename: str | None, fallback: str) -> str:
    if not filename:
        return fallback
    base = Path(filename.replace("\\", "/")).name.strip()
    cleaned = re.sub(r"[\x00-\x1f<>:\"|?*]", "_", base)
    return cleaned[:120] if cleaned else fallback


def detect_format(path: Path) -> str:
    if path.stat().st_size > MAX_FILE_BYTES:
        raise ValueError("El archivo supera el límite de 128 MiB")
    with path.open("rb") as file:
        head = file.read(16)
        if head.startswith(b"%PDF-"):
            return "pdf"
        if head.startswith(b"\x89PNG\r\n\x1a\n"):
            return "png"
        if head.startswith(b"\xff\xd8\xff"):
            return "jpeg"
        if len(head) >= 12 and head[:4] == b"RIFF" and head[8:12] == b"WEBP":
            return "webp"
        file.seek(0)
        try:
            with ZipFile(file) as zip_file:
                entries = zip_file.infolist()
                if len(entries) > MAX_ZIP_ENTRIES or sum(e.file_size for e in entries) > MAX_ZIP_UNCOMPRESSED:
                    raise ValueError("Paquete de documento supera los límites")
                names = set(zip_file.namelist())
                if len(names) != len(entries) or any(n.endswith("vbaProject.bin") for n in names):
                    raise ValueError("Paquete no permitido o contiene macros")
                if "xl/workbook.xml" in names:
                    return "xlsx"
                if "word/document.xml" in names:
                    return "docx"
        except BadZipFile:
            pass
    raise ValueError("Formato no admitido")


def load_case_manifest(root: Path, case_id: UUID) -> dict[str, dict[str, Any]]:
    manifest_path = root / str(case_id) / "metadata.json"
    if not manifest_path.is_file():
        return {}
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
        docs = data.get("documents", {})
        return docs if isinstance(docs, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def _hidden_columns_for_sheet(sheet) -> set[int]:
    hidden: set[int] = set()
    for dimension in sheet.column_dimensions.values():
        if dimension.hidden:
            start = dimension.min or column_index_from_string(dimension.index)
            end = dimension.max or start
            hidden.update(range(start, end + 1))
    return hidden


def _hidden_rows_for_sheet(sheet) -> set[int]:
    return {row_idx for row_idx, dim in sheet.row_dimensions.items() if dim.hidden}


def _extract_formula_text(cell) -> str | None:
    if cell.data_type != "f":
        return None
    if isinstance(cell.value, str):
        return cell.value
    return getattr(cell.value, "text", None)


def _normalize_formula_pattern(formula: str) -> str:
    stripped = _STRING_LITERAL_RE.sub('""', formula)
    return re.sub(r"\d+", "#", stripped)


def _extract_references(formula: str, default_sheet: str, valid_sheets: set[str]) -> list[tuple[str, str]]:
    """Extract referenced (sheet, cell) pairs from an Excel formula. Caps huge ranges cleanly."""
    cleaned = _STRING_LITERAL_RE.sub("", formula)
    cleaned = _EXTERNAL_LINK_RE.sub("", cleaned)
    refs: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for match in _REF_TOKEN_RE.finditer(cleaned):
        quoted_sheet, plain_sheet, c1, r1, c2, r2 = match.groups()
        target_sheet = (
            quoted_sheet.replace("''", "'") if quoted_sheet else (plain_sheet or default_sheet)
        )
        if target_sheet not in valid_sheets:
            continue
        col1 = column_index_from_string(c1)
        row1 = int(r1)
        if c2 and r2:
            col2 = column_index_from_string(c2)
            row2 = int(r2)
            min_c, max_c = sorted((col1, col2))
            min_r, max_r = sorted((row1, row2))
            area = (max_c - min_c + 1) * (max_r - min_r + 1)
            if area <= 400:
                for rr in range(min_r, max_r + 1):
                    for cc in range(min_c, max_c + 1):
                        pair = (target_sheet, f"{get_column_letter(cc)}{rr}")
                        if pair not in seen:
                            seen.add(pair)
                            refs.append(pair)
            else:
                # Sample endpoints for massive full-column ranges
                for pair in (
                    (target_sheet, f"{get_column_letter(min_c)}{min_r}"),
                    (target_sheet, f"{get_column_letter(max_c)}{max_r}"),
                ):
                    if pair not in seen:
                        seen.add(pair)
                        refs.append(pair)
        else:
            pair = (target_sheet, f"{c1}{row1}")
            if pair not in seen:
                seen.add(pair)
                refs.append(pair)
    return refs


def _row_context_label(val_sheet, row_idx: int, max_col: int) -> str:
    parts: list[str] = []
    for c in range(1, min(max_col, 8) + 1):
        v = val_sheet.cell(row=row_idx, column=c).value
        if isinstance(v, str) and v.strip():
            parts.append(v.strip()[:60])
            if len(parts) >= 2:
                break
    return " · ".join(parts)


@dataclass
class LoadedWorkbook:
    wb_formulas: Any
    wb_values: Any
    sheet_names: set[str]
    hidden_sheets: set[str]
    hidden_rows: dict[str, set[int]]
    hidden_cols: dict[str, set[int]]
    merge_map: dict[str, dict[tuple[int, int], str]]


def load_dual_workbook(path: Path) -> LoadedWorkbook:
    with path.open("rb") as f1:
        wb_f = load_workbook(f1, data_only=False, keep_links=False)
    with path.open("rb") as f2:
        wb_v = load_workbook(f2, data_only=True, keep_links=False)
    sheet_names = {s.title for s in wb_f.worksheets}
    hidden_sheets = {s.title for s in wb_f.worksheets if s.sheet_state != "visible"}
    hidden_rows: dict[str, set[int]] = {}
    hidden_cols: dict[str, set[int]] = {}
    merge_map: dict[str, dict[tuple[int, int], str]] = {}
    for sheet in wb_f.worksheets:
        hidden_rows[sheet.title] = _hidden_rows_for_sheet(sheet)
        hidden_cols[sheet.title] = _hidden_columns_for_sheet(sheet)
        merges: dict[tuple[int, int], str] = {}
        for rng in sheet.merged_cells.ranges:
            anchor = f"{get_column_letter(rng.min_col)}{rng.min_row}"
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merges[(r, c)] = anchor
        merge_map[sheet.title] = merges
    return LoadedWorkbook(wb_f, wb_v, sheet_names, hidden_sheets, hidden_rows, hidden_cols, merge_map)


def is_cell_hidden(loaded: LoadedWorkbook, sheet_name: str, coord: str) -> bool:
    if sheet_name in loaded.hidden_sheets:
        return True
    row, col = coordinate_to_tuple(coord)
    return row in loaded.hidden_rows.get(sheet_name, set()) or col in loaded.hidden_cols.get(sheet_name, set())


def resolve_excel_target(path: Path, sheet_name: str, cell_coord: str) -> tuple[str | None, str | None, str | None]:
    """Validate and normalize an Excel target (resolving merged cells to their anchor)."""
    if not sheet_name or not cell_coord or not _CELL_COORD_RE.match(cell_coord):
        return None, None, "Coordenada o nombre de hoja inválido"
    loaded = load_dual_workbook(path)
    if sheet_name not in loaded.sheet_names:
        # Try case-insensitive match
        matched = next((s for s in loaded.sheet_names if s.casefold() == sheet_name.casefold()), None)
        if not matched:
            return None, None, f"Hoja '{sheet_name}' no encontrada"
        sheet_name = matched
    if sheet_name in loaded.hidden_sheets:
        return None, None, f"La hoja '{sheet_name}' está oculta"
    row, col = coordinate_to_tuple(cell_coord)
    anchor = loaded.merge_map.get(sheet_name, {}).get((row, col))
    if anchor:
        cell_coord = anchor
        row, col = coordinate_to_tuple(cell_coord)
    if row in loaded.hidden_rows.get(sheet_name, set()) or col in loaded.hidden_cols.get(sheet_name, set()):
        return None, None, f"La celda {sheet_name}!{cell_coord} está en una fila o columna oculta"
    f_cell = loaded.wb_formulas[sheet_name][cell_coord]
    v_cell = loaded.wb_values[sheet_name][cell_coord]
    formula = _extract_formula_text(f_cell)
    has_val = v_cell.value is not None and str(v_cell.value).strip() != ""
    has_raw = f_cell.value is not None and str(f_cell.value).strip() != ""
    if not formula and not has_val and not has_raw:
        return None, None, f"La celda {sheet_name}!{cell_coord} está vacía"
    if formula and formula.strip() == '=""':
        return None, None, f"La celda {sheet_name}!{cell_coord} devuelve cadena vacía"
    return sheet_name, cell_coord, None


def index_workbook(path: Path, loaded: LoadedWorkbook | None = None) -> dict[str, Any]:
    """Full 100%-coverage structural & technical scan of an Excel workbook of any size."""
    loaded = loaded or load_dual_workbook(path)
    sheets_summary: list[dict[str, Any]] = []
    formula_errors: list[dict[str, Any]] = []
    external_links: list[dict[str, Any]] = []
    hidden_dependencies: list[dict[str, Any]] = []
    formula_breaks: list[dict[str, Any]] = []
    total_visible_cells = 0
    total_formulas = 0

    for sheet in loaded.wb_formulas.worksheets:
        sname = sheet.title
        visible = sheet.sheet_state == "visible"
        if not visible:
            sheets_summary.append({"sheet": sname, "visible": False})
            continue

        val_sheet = loaded.wb_values[sname]
        h_rows = loaded.hidden_rows[sname]
        h_cols = loaded.hidden_cols[sname]
        min_r = min_c = None
        max_r = max_c = 0
        populated = 0
        formulas_count = 0
        headings: list[dict[str, Any]] = []
        col_series: dict[int, list[tuple[int, str, str | None, Any]]] = {}

        for row in sheet.iter_rows():
            if not row:
                continue
            r_idx = row[0].row
            if r_idx in h_rows:
                continue
            row_non_empty = 0
            first_text_cell = None
            for cell in row:
                c_idx = cell.column
                if c_idx in h_cols:
                    continue
                formula = _extract_formula_text(cell)
                cached = val_sheet.cell(row=r_idx, column=c_idx).value
                raw_val = cell.value if not formula else None
                effective = cached if cached is not None else raw_val
                if formula is None and (effective is None or str(effective).strip() == ""):
                    continue

                populated += 1
                row_non_empty += 1
                if min_r is None or r_idx < min_r:
                    min_r = r_idx
                if min_c is None or c_idx < min_c:
                    min_c = c_idx
                if r_idx > max_r:
                    max_r = r_idx
                if c_idx > max_c:
                    max_c = c_idx

                if isinstance(effective, str) and first_text_cell is None:
                    txt = effective.strip()
                    if 3 <= len(txt) <= 90 and not txt.startswith("="):
                        first_text_cell = (cell.coordinate, txt)

                if formula:
                    formulas_count += 1
                    col_series.setdefault(c_idx, []).append((r_idx, "f", _normalize_formula_pattern(formula), formula))
                    if _EXTERNAL_LINK_RE.search(formula) and len(external_links) < 40:
                        external_links.append({
                            "sheet": sname,
                            "cell": cell.coordinate,
                            "formula": clip(formula, 200),
                            "cached_value": clip(cached, 120),
                            "context": _row_context_label(val_sheet, r_idx, max_c),
                        })
                    # Detect hidden precedents for visible formula cells
                    if len(hidden_dependencies) < 40:
                        refs = _extract_references(formula, sname, loaded.sheet_names)
                        hidden_refs = []
                        for ref_s, ref_c in refs:
                            if is_cell_hidden(loaded, ref_s, ref_c):
                                ref_val = loaded.wb_values[ref_s][ref_c].value
                                if ref_val in EXCEL_ERRORS or ref_val is not None:
                                    hidden_refs.append({
                                        "target": f"{ref_s}!{ref_c}",
                                        "value": clip(ref_val, 80),
                                        "is_error": str(ref_val) in EXCEL_ERRORS,
                                    })
                                    if len(hidden_refs) >= 5:
                                        break
                        if any(hr["is_error"] for hr in hidden_refs) or (
                            hidden_refs and (str(cached) in EXCEL_ERRORS or cached == 0)
                        ):
                            hidden_dependencies.append({
                                "sheet": sname,
                                "cell": cell.coordinate,
                                "formula": clip(formula, 200),
                                "cached_value": clip(cached, 120),
                                "hidden_precedents": hidden_refs,
                                "context": _row_context_label(val_sheet, r_idx, max_c),
                            })
                elif isinstance(effective, (int, float)) and not isinstance(effective, bool):
                    col_series.setdefault(c_idx, []).append((r_idx, "n", None, effective))

                if isinstance(cached, str) and cached in EXCEL_ERRORS and len(formula_errors) < 60:
                    formula_errors.append({
                        "sheet": sname,
                        "cell": cell.coordinate,
                        "error": cached,
                        "formula": clip(formula, 200),
                        "context": _row_context_label(val_sheet, r_idx, max_c),
                    })

            # Collect section headings (short header rows or uppercase section titles)
            if first_text_cell and len(headings) < 18:
                coord, txt = first_text_cell
                if row_non_empty <= 3 or txt.isupper():
                    headings.append({"cell": coord, "text": txt[:80]})

        # Detect formula-break anomalies (hardcoded number surrounded by consistent formulas in same column)
        if len(formula_breaks) < 30:
            for c_idx, items in col_series.items():
                if len(items) < 4:
                    continue
                col_letter = get_column_letter(c_idx)
                for idx in range(1, len(items) - 1):
                    r_cur, kind_cur, _, val_cur = items[idx]
                    r_prev, kind_prev, pat_prev, f_prev = items[idx - 1]
                    r_next, kind_next, pat_next, _ = items[idx + 1]
                    if (
                        kind_cur == "n"
                        and kind_prev == "f"
                        and kind_next == "f"
                        and pat_prev == pat_next
                        and (r_cur - r_prev) <= 2
                        and (r_next - r_cur) <= 2
                    ):
                        formula_breaks.append({
                            "sheet": sname,
                            "cell": f"{col_letter}{r_cur}",
                            "hardcoded_value": val_cur,
                            "neighbor_formula": clip(f_prev, 120),
                            "context": _row_context_label(val_sheet, r_cur, max_c),
                        })
                        if len(formula_breaks) >= 30:
                            break

        used_range = (
            f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
            if min_r is not None and min_c is not None
            else None
        )
        total_visible_cells += populated
        total_formulas += formulas_count
        sheets_summary.append({
            "sheet": sname,
            "visible": True,
            "used_range": used_range,
            "visible_cells": populated,
            "formula_cells": formulas_count,
            "merged_ranges": len(sheet.merged_cells.ranges),
            "headings": headings,
        })

    return {
        "visible_sheets": sum(1 for s in sheets_summary if s["visible"]),
        "hidden_sheets": sum(1 for s in sheets_summary if not s["visible"]),
        "total_visible_cells": total_visible_cells,
        "total_formulas": total_formulas,
        "sheets": sheets_summary,
        "deterministic_alerts": {
            "formula_errors": formula_errors,
            "external_links": external_links,
            "hidden_dependencies": hidden_dependencies,
            "formula_breaks": formula_breaks,
        },
    }


def search_workbook(
    loaded: LoadedWorkbook,
    query: str,
    sheet_name: str | None = None,
    limit: int = 40,
) -> dict[str, Any]:
    needle = query.strip().casefold()
    if not needle:
        raise ValueError("Indique un término de búsqueda")
    matches: list[dict[str, Any]] = []
    truncated = False
    target_sheets = [loaded.wb_formulas[sheet_name]] if sheet_name else list(loaded.wb_formulas.worksheets)
    for sheet in target_sheets:
        if sheet.sheet_state != "visible":
            if sheet_name:
                raise ValueError("La hoja está oculta")
            continue
        sname = sheet.title
        val_sheet = loaded.wb_values[sname]
        h_rows = loaded.hidden_rows[sname]
        h_cols = loaded.hidden_cols[sname]
        for row in sheet.iter_rows():
            if not row:
                continue
            r_idx = row[0].row
            if r_idx in h_rows:
                continue
            for cell in row:
                if cell.column in h_cols:
                    continue
                formula = _extract_formula_text(cell)
                cached = val_sheet.cell(row=r_idx, column=cell.column).value
                raw = cell.value if not formula else None
                haystack = " ".join(
                    str(x) for x in (formula, cached, raw) if x is not None
                ).casefold()
                if needle in haystack:
                    matches.append({
                        "sheet": sname,
                        "cell": cell.coordinate,
                        "formula": clip(formula, 200),
                        "cached_value": clip(cached if cached is not None else raw, 200),
                        "number_format": clip(cell.number_format, 60),
                        "row_context": _row_context_label(val_sheet, r_idx, sheet.max_column or 8),
                    })
                    if len(matches) >= limit:
                        truncated = True
                        break
            if truncated:
                break
        if truncated:
            break
    return {"query": query, "count": len(matches), "truncated": truncated, "matches": matches}


def trace_cell_dependencies(loaded: LoadedWorkbook, sheet_name: str, cell_coord: str) -> dict[str, Any]:
    if sheet_name not in loaded.sheet_names:
        raise ValueError("Hoja no encontrada")
    if sheet_name in loaded.hidden_sheets:
        raise ValueError("La hoja está oculta")
    cell_coord = cell_coord.strip().upper()
    if not _CELL_COORD_RE.match(cell_coord):
        raise ValueError("Coordenada inválida")
    row, col = coordinate_to_tuple(cell_coord)
    anchor = loaded.merge_map.get(sheet_name, {}).get((row, col))
    if anchor:
        cell_coord = anchor
    if is_cell_hidden(loaded, sheet_name, cell_coord):
        raise ValueError("La celda está oculta")

    f_cell = loaded.wb_formulas[sheet_name][cell_coord]
    v_cell = loaded.wb_values[sheet_name][cell_coord]
    formula = _extract_formula_text(f_cell)

    precedents: list[dict[str, Any]] = []
    if formula:
        for ref_s, ref_c in _extract_references(formula, sheet_name, loaded.sheet_names)[:60]:
            rf_cell = loaded.wb_formulas[ref_s][ref_c]
            rv_cell = loaded.wb_values[ref_s][ref_c]
            precedents.append({
                "sheet": ref_s,
                "cell": ref_c,
                "hidden": is_cell_hidden(loaded, ref_s, ref_c),
                "formula": clip(_extract_formula_text(rf_cell), 200),
                "cached_value": clip(rv_cell.value, 200),
            })

    # Find visible cells across the workbook that directly reference (sheet_name, cell_coord)
    dependents: list[dict[str, Any]] = []
    target_pair = (sheet_name, cell_coord)
    for sheet in loaded.wb_formulas.worksheets:
        if sheet.sheet_state != "visible":
            continue
        sname = sheet.title
        h_rows = loaded.hidden_rows[sname]
        h_cols = loaded.hidden_cols[sname]
        val_sheet = loaded.wb_values[sname]
        for r in sheet.iter_rows():
            if not r or r[0].row in h_rows:
                continue
            for c in r:
                if c.column in h_cols or c.data_type != "f":
                    continue
                cand_f = _extract_formula_text(c)
                if not cand_f:
                    continue
                # Fast pre-check before full regex parsing
                if cell_coord not in cand_f and ":" not in cand_f:
                    continue
                if target_pair in _extract_references(cand_f, sname, loaded.sheet_names):
                    dependents.append({
                        "sheet": sname,
                        "cell": c.coordinate,
                        "formula": clip(cand_f, 200),
                        "cached_value": clip(val_sheet[c.coordinate].value, 200),
                        "row_context": _row_context_label(val_sheet, c.row, sheet.max_column or 8),
                    })
                    if len(dependents) >= 50:
                        break
            if len(dependents) >= 50:
                break
        if len(dependents) >= 50:
            break

    return {
        "target": {
            "sheet": sheet_name,
            "cell": cell_coord,
            "formula": clip(formula, 250),
            "cached_value": clip(v_cell.value, 250),
            "number_format": clip(f_cell.number_format, 100),
            "has_external_link": bool(formula and _EXTERNAL_LINK_RE.search(formula)),
        },
        "precedents": precedents,
        "visible_dependents": dependents,
    }


def pdftoppm_available() -> bool:
    return shutil.which("pdftoppm") is not None


def render_pdf_page_png(path: Path, page: int, dpi: int = 130) -> bytes:
    exe = shutil.which("pdftoppm")
    if not exe:
        raise ValueError("Renderizado visual de PDF no disponible (falta pdftoppm)")
    dpi = max(72, min(dpi, 200))
    with TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        prefix = Path(tmp) / "page"
        try:
            subprocess.run(
                [exe, "-png", "-r", str(dpi), "-f", str(page), "-l", str(page), str(path), str(prefix)],
                check=True,
                timeout=25,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except (OSError, subprocess.SubprocessError):
            raise ValueError(f"No se pudo renderizar visualmente la página {page} del PDF") from None
        rendered = sorted(Path(tmp).glob("page*.png"))
        if not rendered:
            raise ValueError(f"No se generó imagen para la página {page}")
        return rendered[0].read_bytes()


def index_pdf(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        reader = PdfReader(stream)
        if reader.is_encrypted:
            raise ValueError("El PDF está cifrado y no se puede leer")
        page_count = len(reader.pages)
        searchable_pages: list[int] = []
        scanned_pages: list[int] = []
        outline: list[dict[str, Any]] = []
        for idx, page in enumerate(reader.pages, start=1):
            text = (page.extract_text() or "").strip()
            has_text = len(text) >= 15
            if has_text:
                searchable_pages.append(idx)
            else:
                scanned_pages.append(idx)
            if len(outline) < 60:
                first_line = next((ln.strip() for ln in text.splitlines() if ln.strip()), "")
                outline.append({
                    "page": idx,
                    "searchable": has_text,
                    "chars": len(text),
                    "preview": first_line[:110] if first_line else "(sin texto extraíble / escaneada)",
                })
        return {
            "page_count": page_count,
            "searchable_page_count": len(searchable_pages),
            "scanned_pages": scanned_pages[:100],
            "visual_inspection_available": pdftoppm_available(),
            "outline": outline,
        }


def search_pdf(path: Path, query: str, limit: int = 40) -> dict[str, Any]:
    needle = query.strip().casefold()
    if not needle:
        raise ValueError("Indique un término de búsqueda")
    matches: list[dict[str, Any]] = []
    truncated = False
    with path.open("rb") as stream:
        reader = PdfReader(stream)
        if reader.is_encrypted:
            raise ValueError("El PDF está cifrado")
        for idx, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            lowered = text.casefold()
            pos = lowered.find(needle)
            if pos != -1:
                start = max(0, pos - 120)
                end = min(len(text), pos + len(needle) + 180)
                snippet = re.sub(r"\s+", " ", text[start:end]).strip()
                matches.append({"page": idx, "snippet": snippet})
                if len(matches) >= limit:
                    truncated = True
                    break
    return {"query": query, "count": len(matches), "truncated": truncated, "matches": matches}


def index_docx(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        doc = Document(stream)
        blocks = list(doc.iter_inner_content())
        paragraphs = 0
        tables = 0
        previews: list[dict[str, Any]] = []
        for idx, block in enumerate(blocks, start=1):
            if hasattr(block, "rows"):
                tables += 1
                first_row = " | ".join(c.text.strip() for c in block.rows[0].cells)[:120] if block.rows else ""
                if len(previews) < 40:
                    previews.append({"block": idx, "type": "table", "rows": len(block.rows), "preview": first_row})
            else:
                paragraphs += 1
                txt = block.text.strip()
                if txt and len(previews) < 40:
                    previews.append({"block": idx, "type": "paragraph", "preview": txt[:120]})
        return {
            "block_count": len(blocks),
            "paragraphs": paragraphs,
            "tables": tables,
            "previews": previews,
        }


def search_docx(path: Path, query: str, limit: int = 40) -> dict[str, Any]:
    needle = query.strip().casefold()
    if not needle:
        raise ValueError("Indique un término de búsqueda")
    matches: list[dict[str, Any]] = []
    truncated = False
    with path.open("rb") as stream:
        doc = Document(stream)
        for idx, block in enumerate(doc.iter_inner_content(), start=1):
            if hasattr(block, "rows"):
                for r_idx, row in enumerate(block.rows, start=1):
                    text = " | ".join(cell.text.strip() for cell in row.cells)
                    if needle in text.casefold():
                        matches.append({"block": idx, "type": "table", "row": r_idx, "snippet": text[:300]})
                        if len(matches) >= limit:
                            truncated = True
                            break
            else:
                text = block.text.strip()
                if needle in text.casefold():
                    matches.append({"block": idx, "type": "paragraph", "snippet": text[:300]})
                    if len(matches) >= limit:
                        truncated = True
            if truncated:
                break
        if not truncated:
            for s_idx, section in enumerate(doc.sections, start=1):
                for zone, container in (("header", section.header), ("footer", section.footer)):
                    if not container:
                        continue
                    txt = " ".join(p.text.strip() for p in container.paragraphs if p.text.strip())
                    if txt and needle in txt.casefold():
                        matches.append({"block": 0, "type": f"section_{s_idx}_{zone}", "snippet": txt[:300]})
                        if len(matches) >= limit:
                            truncated = True
                            break
                if truncated:
                    break
    return {"query": query, "count": len(matches), "truncated": truncated, "matches": matches}


def build_case_atlas(root: Path, case_id: UUID) -> dict[str, Any]:
    """Build complete deterministic structural atlas for all documents in a case."""
    root = root.resolve(strict=True)
    originals = root / str(case_id) / "originals"
    manifest = load_case_manifest(root, case_id)
    documents: list[dict[str, Any]] = []
    coverage_notes: list[str] = []

    ordered_ids = list(manifest.keys())
    for path in sorted(originals.iterdir(), key=lambda p: p.name):
        if path.name not in ordered_ids:
            try:
                if str(UUID(path.name)) == path.name:
                    ordered_ids.append(path.name)
            except ValueError:
                continue

    for doc_id in ordered_ids:
        path = originals / doc_id
        if not path.is_file():
            continue
        meta = manifest.get(doc_id, {})
        role = meta.get("role", "documento")
        label = meta.get("label", doc_id)
        try:
            fmt = detect_format(path)
        except ValueError as exc:
            documents.append({"document_id": doc_id, "role": role, "label": label, "format": "no_admitido", "error": str(exc)})
            coverage_notes.append(f"{label}: formato no admitido ({exc}).")
            continue

        entry: dict[str, Any] = {
            "document_id": doc_id,
            "role": role,
            "label": label,
            "format": fmt,
            "size_bytes": path.stat().st_size,
        }
        try:
            if fmt == "xlsx":
                info = index_workbook(path)
                entry["summary"] = info
                coverage_notes.append(
                    f"{label} (Excel {role}): {info['visible_sheets']} hojas visibles "
                    f"({info['hidden_sheets']} ocultas omitidas), {info['total_visible_cells']} celdas con datos, "
                    f"{info['total_formulas']} fórmulas escaneadas al 100%."
                )
            elif fmt == "pdf":
                info = index_pdf(path)
                entry["summary"] = info
                scanned_count = len(info["scanned_pages"])
                extra = (
                    f"{scanned_count} páginas escaneadas/sin texto (inspección visual {'disponible' if info['visual_inspection_available'] else 'no disponible'})"
                    if scanned_count
                    else "texto extraíble en todas las páginas"
                )
                coverage_notes.append(f"{label} (PDF {role}): {info['page_count']} páginas; {extra}.")
            elif fmt == "docx":
                info = index_docx(path)
                entry["summary"] = info
                coverage_notes.append(
                    f"{label} (Word {role}): {info['paragraphs']} párrafos y {info['tables']} tablas indexados."
                )
            elif fmt in IMAGE_FORMATS:
                entry["summary"] = {"mime_type": IMAGE_FORMATS[fmt], "visual_inspection_available": True}
                coverage_notes.append(f"{label} (Imagen {role}): disponible para inspección visual directa.")
        except Exception as exc:
            entry["error"] = str(exc)
            coverage_notes.append(f"{label}: no pudo indexarse ({exc}).")
        documents.append(entry)

    return {"case_id": str(case_id), "documents": documents, "coverage": coverage_notes}
