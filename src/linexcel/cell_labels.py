"""Bounded, deterministic workbook labels, independent of any model.

Labels report workbook text and geometry, never a verified business meaning.
Metadata rectangles use a spatial index; a lookup never enumerates sheet cells.
"""

from __future__ import annotations

import io
import sqlite3
import zipfile
from collections.abc import Callable, Iterable
from typing import Any

from openpyxl.utils.cell import coordinate_to_tuple

from linexcel.refs import a1, parse_ref

SCAN_DISTANCE = 8
MAX_LABEL_CHARS = 240
MAX_LABELS = 16


def _text(value: Any) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None
    try:
        float(value.strip().replace(",", "."))
        return None  # Numeric-looking text is also a doubtful heading.
    except ValueError:
        return value


class CellLabelIndex:
    """Index metadata once; probe at most eight cells in each direction.

    ``cell`` returns literal source text only, excluding formulas and errors.
    Tables use the declared OOXML schema (bounds, headers, headerRows).
    Names are exact single-cell references; formulas/ranges stay unresolved.
    """

    def __init__(
        self,
        cell: Callable[[str, int, int], Any],
        *,
        tables: Iterable[dict] = (),
        merges: Iterable[tuple[str, tuple[int, int, int, int]]] = (),
        names: Iterable[dict] = (),
    ):
        self.cell = cell
        self.db = sqlite3.connect(":memory:")
        self.db.execute("CREATE VIRTUAL TABLE boxes USING rtree(id,s1,s2,r1,r2,c1,c2)")
        self.sheets: dict[str, int] = {}
        self.records: dict[int, tuple[str, str, Any]] = {}
        records = (
            (("merge", sheet, bounds, (sheet, bounds)) for sheet, bounds in merges),
            (("table", t["sheet"], t["bounds"], t) for t in tables),
        )
        for entries in records:
            for kind, sheet, bounds, record in entries:
                r1, c1, r2, c2 = bounds
                key = len(self.records)
                sheet_id = self.sheets.setdefault(sheet, len(self.sheets))
                self.records[key] = (kind, sheet, record)
                self.db.execute(
                    "INSERT INTO boxes VALUES (?,?,?,?,?,?,?)",
                    (key, sheet_id, sheet_id, r1, r2, c1, c2),
                )
        self.names: dict[tuple, list[dict]] = {}
        for name in names:
            expression = name.get("expression")
            if not isinstance(expression, str):
                continue
            rect = parse_ref(expression.removeprefix("="))
            if rect is None or (rect.r1, rect.c1) != (rect.r2, rect.c2):
                continue
            sheet = rect.sheet or name.get("scope_sheet")
            if sheet and name.get("scope") in {"workbook", "sheet"}:
                self.names.setdefault((sheet, rect.r1, rect.c1), []).append(name)

    def close(self) -> None:
        self.db.close()

    def boxes(self, sheet: str, row: int, column: int, kind: str) -> list:
        hits = []
        for (key,) in self.db.execute(
            "SELECT id FROM boxes WHERE s1=? AND s2=? "
            "AND r1<=? AND r2>=? AND c1<=? AND c2>=? LIMIT 33",
            (
                self.sheets.get(sheet, -1),
                self.sheets.get(sheet, -1),
                row,
                row,
                column,
                column,
            ),
        ):
            record_kind, record_sheet, record = self.records[key]
            if record_kind == kind and record_sheet == sheet:
                hits.append(record)
                if len(hits) > MAX_LABELS:
                    return []  # Ambiguous overlapping metadata.
        return hits

    def anchor(self, sheet: str, row: int, column: int) -> tuple[int, int] | None:
        merges = self.boxes(sheet, row, column, "merge")
        if len(merges) > 1:
            return None
        return tuple(merges[0][1][:2]) if merges else (row, column)

    def labels(self, sheet: str, row: int, column: int) -> list[dict]:
        result: list[dict] = []
        target_anchor = self.anchor(sheet, row, column)

        def add(text, r, c, method, role, **extra):
            if (text := _text(text)) is not None and len(result) < MAX_LABELS:
                result.append(
                    dict(
                        text=text[:MAX_LABEL_CHARS],
                        sheet=sheet,
                        cell=a1(r, c),
                        method=method,
                        role=role,
                        truncated=len(text) > MAX_LABEL_CHARS,
                        **extra,
                    )
                )

        for table in self.boxes(sheet, row, column, "table"):
            if table.get("headerRows") != 1:
                continue  # No source header cell for a headerless table.
            r1, c1, _, _ = table["bounds"]
            headers = table["headers"]
            if column - c1 < len(headers):
                add(headers[column - c1], r1, column, "table_column", "column")
        for name in self.names.get((sheet, row, column), [])[:MAX_LABELS]:
            add(
                name.get("name"),
                row,
                column,
                "defined_name",
                "name",
                scope=name.get("scope"),
                scope_sheet=name.get("scope_sheet"),
            )
        for method, role, dr, dc in (
            ("left_text", "row", 0, -1),
            ("above_text", "column", -1, 0),
        ):
            for distance in range(1, SCAN_DISTANCE + 1):
                r, c = row + dr * distance, column + dc * distance
                if min(r, c) < 1:
                    break
                anchor = self.anchor(sheet, r, c)
                if anchor is None or anchor == target_anchor:
                    continue
                r, c = anchor
                # Merges can move the anchor off the scan axis, but cannot
                # reach arbitrarily far away or claim the target's own text.
                if max(abs(row - r), abs(column - c)) > SCAN_DISTANCE:
                    continue
                value = self.cell(sheet, r, c)
                if _text(value) is not None:
                    add(value, r, c, method, role)
                    break
        return result


def enrich_graph_labels(data: bytes, nodes: Iterable[dict]) -> None:
    """One sparse source pass per sheet for a batch; no scan per node.

    Only the bounded probe cells are retained, and only referenced shared
    strings are decoded. Formula caches cannot become labels.
    """
    from linexcel.doc_evidence import read_defined_name_evidence
    from linexcel.formula_source import _cells, _worksheet_parts
    from linexcel.lazy_store import _decode_xstring, _stream_elements
    from linexcel.tables import _declared_tables

    targets = []
    for node in nodes:
        if node.get("kind") not in {"cell", "group", "input"}:
            continue
        sheet, address = node.get("sheet"), node.get("addr", node.get("cell"))
        if not sheet or not isinstance(address, str):
            continue
        try:
            row, column = coordinate_to_tuple(address)
        except ValueError:
            continue
        targets.append((node, sheet, row, column))
    if not targets:
        return
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = dict(_worksheet_parts(archive))
        merges = []
        for sheet, part in parts.items():
            for element in _stream_elements(archive, part, "mergeCell"):
                rect = parse_ref(element.get("ref", ""))
                if rect:
                    merges.append((sheet, (rect.r1, rect.c1, rect.r2, rect.c2)))
        tables = [
            dict(
                sheet=sheet,
                bounds=(t["header_row"], t["first_col"], t["last_row"], t["last_col"]),
                headers=t["headers"],
                headerRows=t["first_row"] - t["header_row"],
            )
            for sheet, entries in _declared_tables(data).items()
            for t in entries
        ]
        cells: dict[tuple, Any] = {}
        index = CellLabelIndex(
            lambda s, r, c: cells.get((s, r, c)),
            merges=merges,
            tables=tables,
            names=read_defined_name_evidence(data).get("definitions", []),
        )
        try:
            probes: dict[str, set[str]] = {}
            for _, sheet, row, column in targets:
                wanted = probes.setdefault(sheet, set())
                for dr, dc in ((0, -1), (-1, 0)):
                    for distance in range(1, SCAN_DISTANCE + 1):
                        r, c = row + dr * distance, column + dc * distance
                        if min(r, c) > 0:
                            anchor = index.anchor(sheet, r, c)
                            if anchor:
                                wanted.add(a1(*anchor))
            shared: dict[int, list[tuple]] = {}
            for sheet, wanted in probes.items():
                if sheet not in parts:
                    continue
                for element, address in _cells(archive, parts[sheet]):
                    if address not in wanted or element.find("{*}f") is not None:
                        continue
                    key = (sheet, *coordinate_to_tuple(address))
                    kind = element.get("t")
                    if kind == "s":
                        try:
                            shared.setdefault(
                                int(element.findtext("{*}v", "")), []
                            ).append(key)
                        except ValueError:
                            pass
                    elif kind in {"inlineStr", "str"}:
                        cells[key] = _decode_xstring(
                            "".join(
                                t.text or ""
                                for t in element.findall("{*}is/{*}t")
                                + element.findall("{*}is/{*}r/{*}t")
                            )
                            if kind == "inlineStr"
                            else element.findtext("{*}v", "")
                        )
            if shared and "xl/sharedStrings.xml" in archive.namelist():
                for number, element in enumerate(
                    _stream_elements(archive, "xl/sharedStrings.xml", "si")
                ):
                    if number in shared:
                        text = _decode_xstring(
                            "".join(
                                t.text or ""
                                for t in element.findall("{*}t")
                                + element.findall("{*}r/{*}t")
                            )
                        )
                        for key in shared[number]:
                            cells[key] = text
            for node, sheet, row, column in targets:
                node["cellLabels"] = index.labels(sheet, row, column)
        finally:
            index.close()
