"""Workbook tables and defined names, read from the package parts.

Extracted mechanically from analyzer.py. Declared Excel tables come from the
package's own ``xl/tables/*.xml`` parts (via the relationship graph of the
OOXML format), static ones from a heuristic over a small engine window, and
defined names from openpyxl's read-only workbook. All of it is used only to
tag lineage nodes — nothing here builds the graph itself.
"""

from __future__ import annotations

import io
import zipfile
from collections import defaultdict
from typing import Any
from xml.etree import ElementTree

import formualizer as fz

from linexcel.refs import Rect, parse_ref, parse_ref_detailed


def _build_table_index(
    data: bytes,
    engine,
    sheet_dims: dict[str, tuple[int, int]],
    engine_sheets: set[str],
) -> dict[str, list[dict[str, Any]]]:
    """Per-sheet table list, for cell→table enrichment of lineage nodes.

    Two sources, both cheap. Declared tables come from the package's own
    ``xl/tables/*.xml`` parts, which carry the range and the column names that
    structured references (``Sales[Amount]``) resolve against. Static ones stay
    a heuristic, over a window of at most 30 × 50 cells read from the engine.

    Neither needs openpyxl. Reaching ``ws.tables`` did — read-only mode drops
    them — and that second, eager parse of the whole workbook cost 45 s of a
    246 s run on a 2.6M-formula file, to read a handful of definitions.
    """
    from linexcel.insights import (
        MAX_TABLES_PER_SHEET,
        STATIC_TABLE_SCAN_COLS,
        STATIC_TABLE_SCAN_ROWS,
        static_tables_from_rows,
    )

    index: dict[str, list[dict[str, Any]]] = {}
    declared = _declared_tables(data)
    for sheet, (max_row, max_col) in sheet_dims.items():
        tables = declared.get(sheet, [])[:MAX_TABLES_PER_SHEET]
        covered = [
            (t["header_row"], t["first_col"], t["last_row"], t["last_col"])
            for t in tables
        ]
        if sheet in engine_sheets:
            rows = _sheet_window(
                engine,
                sheet,
                min(max_row, STATIC_TABLE_SCAN_ROWS),
                min(max_col, STATIC_TABLE_SCAN_COLS),
            )
            try:
                tables = tables + static_tables_from_rows(rows, covered)
            except Exception:
                pass
        if tables:
            index[sheet] = tables
    return index


def _sheet_window(engine, sheet: str, rows: int, cols: int) -> list[list[Any]]:
    """Top-left window of a sheet, as ``data_only=False`` would hand it over.

    A formula cell reads back as its ``=…`` text rather than as its result:
    the static-table heuristic tells a text header from a computed one that
    way, so the window has to keep the distinction openpyxl gave it.
    """
    if rows < 1 or cols < 1:
        return []
    try:
        fsheet = engine.sheet(sheet)
        address = fz.RangeAddress(sheet, 1, 1, rows, cols)
        values = fsheet.get_values(address)
        formulas = fsheet.get_formulas(address)
    except Exception:
        return []
    window: list[list[Any]] = []
    for r, row in enumerate(values):
        frow = formulas[r] if r < len(formulas) else ()
        line: list[Any] = []
        for c, value in enumerate(row):
            formula = frow[c] if c < len(frow) else ""
            if formula:
                line.append(formula if formula.startswith("=") else "=" + formula)
            else:
                line.append(value if value != "" else None)
        window.append(line)
    return window


def _declared_tables(data: bytes) -> dict[str, list[dict[str, Any]]]:
    """Table definitions per sheet, read straight from the package parts.

    Follows the relationships the format itself uses: workbook → worksheet
    parts → the table parts each sheet owns. Anything unreadable is skipped —
    a missing table only costs a node its ``table_name`` tag.
    """
    out: dict[str, list[dict[str, Any]]] = defaultdict(list)
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            book = _xml_part(zf, "xl/workbook.xml")
            if book is None:
                return {}
            rels = _part_rels(zf, "xl/workbook.xml")
            for element in book.iter():
                if _local_name(element.tag) != "sheet":
                    continue
                name = element.get("name")
                rid = next(
                    (v for k, v in element.attrib.items() if _local_name(k) == "id"),
                    None,
                )
                part = rels.get(rid or "")
                if not name or not part:
                    continue
                for table_part in _part_rels(zf, part, kind="table").values():
                    entry = _table_entry(_xml_part(zf, table_part))
                    if entry is not None:
                        out[name].append(entry)
    except Exception:
        return {}
    return dict(out)


def _table_entry(root) -> dict[str, Any] | None:
    """One ``xl/tables/tableN.xml`` part as the index's table dict."""
    if root is None:
        return None
    rect = parse_ref(root.get("ref") or "")
    if rect is None:
        return None
    # An absent headerRowCount means one header row; a table declared without
    # headers says so with a 0, and its first row is data.
    try:
        header_rows = int(root.get("headerRowCount", "1"))
    except ValueError:
        header_rows = 1
    headers = [
        column.get("name")
        for column in root.iter()
        if _local_name(column.tag) == "tableColumn"
    ]
    first_row = rect.r1 + max(header_rows, 0)
    return {
        "name": root.get("displayName") or root.get("name") or "Table",
        "kind": "dynamic",
        "ref": rect.to_a1(),
        "header_row": rect.r1,
        "first_row": first_row,
        "last_row": rect.r2,
        "first_col": rect.c1,
        "last_col": rect.c2,
        "headers": headers,
        "data_rows": max(0, rect.r2 - first_row + 1),
    }


def _local_name(tag: str) -> str:
    """``{namespace}sheet`` → ``sheet``."""
    return tag.rsplit("}", 1)[-1]


def _xml_part(zf: zipfile.ZipFile, part: str):
    try:
        return ElementTree.fromstring(zf.read(part))
    except (KeyError, ElementTree.ParseError):
        return None


def _part_rels(
    zf: zipfile.ZipFile, part: str, kind: str | None = None
) -> dict[str, str]:
    """``{relationship id: part path}`` for one part, optionally by type."""
    base, _, name = part.rpartition("/")
    root = _xml_part(zf, f"{base}/_rels/{name}.rels")
    if root is None:
        return {}
    out: dict[str, str] = {}
    for rel in root:
        rid, target = rel.get("Id"), rel.get("Target")
        if not rid or not target:
            continue
        if kind is not None and not rel.get("Type", "").endswith("/" + kind):
            continue
        out[rid] = _resolve_part(base, target)
    return out


def _resolve_part(base: str, target: str) -> str:
    """Resolve a relationship target against the part that declares it."""
    if target.startswith("/"):
        return target.lstrip("/")
    segments = base.split("/") if base else []
    for segment in target.split("/"):
        if segment == "..":
            if segments:
                segments.pop()
        elif segment not in ("", "."):
            segments.append(segment)
    return "/".join(segments)


def _enrich_with_table(
    node: dict[str, Any],
    table_index: dict[str, list[dict[str, Any]]],
    sheet: str,
    row: int,
    col: int,
) -> None:
    """Tag a node with ``table_name``/``table_column``/``table_row``.

    The representative cell (``row``, ``col``) is checked against every table
    on its sheet; the first match wins. ``table_column`` is the header text of
    the column the cell sits in, ``table_row`` the 0-based offset into the
    table's data area (header row excluded).
    """
    for tbl in table_index.get(sheet, ()):
        if (
            tbl["first_col"] <= col <= tbl["last_col"]
            and tbl["header_row"] <= row <= tbl["last_row"]
        ):
            node["table_name"] = tbl["name"]
            node["table_kind"] = tbl["kind"]
            idx = col - tbl["first_col"]
            headers = tbl["headers"]
            node["table_column"] = headers[idx] if 0 <= idx < len(headers) else None
            node["table_row"] = max(0, row - tbl["first_row"])
            return


def _force_dimensions(ws) -> tuple[int, int]:
    max_row = max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
        if max_row > 200_000:
            break
    return max_row or 1, max_col or 1


def _collect_defined_names(owb) -> dict[str, list[Rect]]:
    """Defined names, workbook-scoped first then sheet-scoped.

    A name can be declared on a single worksheet, in which case it lives on that
    sheet rather than on the workbook and is invisible to ``owb.defined_names``.
    Such names are common in real files — a per-sheet ``Total`` or ``Limit`` —
    and skipping them turned every formula using one into an unresolved
    reference.

    Scope is not modelled on the node itself: the graph keys names by their text,
    so a workbook-scoped name wins over a sheet-scoped one of the same name, and
    the first sheet to declare it wins among sheets. That matches how the rest of
    the graph resolves names (case-insensitively, by label) and keeps a shadowed
    name from silently replacing the one most formulas mean.
    """
    out: dict[str, list[Rect]] = {}
    for name, dn in _iter_defined_names(owb):
        if name.startswith("_xlnm.") or name in out:
            continue
        rects: list[Rect] = []
        try:
            for sheet, coord in dn.destinations:
                # openpyxl strips the surrounding quotes but leaves the doubled
                # apostrophes of the escaped form, so a sheet called O'Brien
                # arrives as O''Brien. Storing that verbatim makes to_a1() quote
                # it a second time, and the name never resolves to its sheet.
                sheet = sheet.replace("''", "'") if sheet else sheet
                detail = parse_ref_detailed(coord, default_sheet=sheet)
                if detail is not None:
                    rect = detail.rect
                    rects.append(
                        Rect(sheet or rect.sheet, rect.r1, rect.c1, rect.r2, rect.c2)
                    )
        except Exception:
            continue
        if rects:
            out[name] = rects
    return out


def _iter_defined_names(owb):
    """``(name, DefinedName)`` pairs, workbook scope before sheet scope."""
    try:
        yield from owb.defined_names.items()
    except Exception:
        pass
    for worksheet in getattr(owb, "worksheets", []):
        try:
            yield from worksheet.defined_names.items()
        except Exception:
            continue
