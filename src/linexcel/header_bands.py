"""Conservative, generation-local header inference for library search only.

Only literal source strings become headings. Rectangles end at observed blank,
missing-cell and section boundaries; a scan cap never licenses extrapolation.
The dossier's independent eight-cell label policy is deliberately unchanged.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass
from itertools import groupby
from typing import Any

from linexcel.lazy_store import _unpack
from linexcel.refs import a1


def descriptive(text):
    if not isinstance(text, str) or not text.strip():
        return False
    try:
        float(text.strip().replace(",", "."))
        return False
    except ValueError:
        return True


def literal(node):
    value = node.get("value")
    if (
        not node.get("formula")
        and node.get("sourceDataType") in {"s", "str", "inlineStr"}
        and isinstance(value, str)
        and value.strip()
    ):
        return value
    return None


def numeric(node):
    value = node.get("value")
    return type(value) in {int, float} and math.isfinite(value)


def catalog_labels(index, sheet, row, column, inferred):
    """Apply catalog evidence rules without changing dossier label semantics."""
    original = index.cell_labels(sheet, row, column)
    declared = any(label["method"] == "table_column" for label in original)
    result = [
        label
        for label in original
        if label["method"] != "above_text"
        and not (inferred and label["method"] == "left_text")
    ]
    if declared:
        result = [
            {**label, "role": "row_context"}
            if label["method"] == "left_text"
            else label
            for label in result
        ]
    sources = {(label["sheet"], label["cell"], label["role"]) for label in result}
    for label in inferred:
        if label["role"] == "column" and any(
            existing["method"] == "table_column" for existing in result
        ):
            continue
        key = (label["sheet"], label["cell"], label["role"])
        if key not in sources:
            result.append(label)
            sources.add(key)
    return result


@dataclass
class Band:
    header_row: int
    first_row: int
    last_row: int
    # Exact literal source strings, keyed by their source column.
    headers: dict[int, str]
    row_columns: list[int]


class HeaderBands:
    def __init__(self, index, sheet, tables, limits, expired):
        self.sheet = sheet
        self.bands: list[Band] = []
        self.rows: dict[int, dict[int, dict]] = {}
        self.coverage: dict[str, Any] = {
            "cellsScanned": 0,
            "rowsScanned": 0,
            "candidates": 0,
            "bands": 0,
            "ambiguousCandidates": 0,
            "capsHit": [],
            "complete": True,
            "scope": "outside declared tables with descriptive headers",
        }
        caps = set()
        if not index.db.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' "
            "AND name='merge_definitions'"
        ).fetchone():
            # An older index cannot prove the absence of merged headers or
            # section boundaries. Keep its conservative labels, infer no bands.
            self.coverage.update(
                capsHit=["header_merge_metadata_unavailable"], complete=False
            )
            return
        if len(tables) > limits.header_regions:
            self.coverage.update(capsHit=["header_regions"], complete=False)
            return
        # Ordinary declared columns already have a stronger label. Do not spend
        # inference budgets scanning millions of their body cells again.
        clauses = ["sheet=?", "row IS NOT NULL", "col IS NOT NULL"]
        args = [sheet]
        for table in tables:
            headers = table.get("headers", [])
            if (
                table.get("headerRows") == 1
                and headers
                and all(descriptive(h) for h in headers)
            ):
                r1, c1, r2, c2 = table["bounds"]
                clauses.append("NOT (row BETWEEN ? AND ? AND col BETWEEN ? AND ?)")
                args.extend((r1, r2, c1, c2))
        cursor = index.db.execute(
            "SELECT row,col,data FROM nodes WHERE "
            + " AND ".join(clauses)
            + " AND kind IN ('input','cell','array_result') ORDER BY row,col LIMIT ?",
            (*args, limits.header_scan_cells + 1),
        )
        count = 0
        for row, cells in groupby(cursor, key=lambda record: record[0]):
            if expired():
                caps.add("build_seconds")
                break
            if len(self.rows) >= limits.header_candidate_rows:
                caps.add("header_candidate_rows")
                break
            values = {}
            truncated = False
            for _, column, blob in cells:
                count += 1
                if count > limits.header_scan_cells:
                    caps.add("header_scan_cells")
                    truncated = True
                    break
                if column > limits.header_columns:
                    caps.add("header_columns")
                    truncated = True
                    continue
                values[column] = _unpack(blob)
            if not truncated:
                self.rows[row] = values
            if count > limits.header_scan_cells:
                break
        self.coverage.update(cellsScanned=min(count, limits.header_scan_cells))
        merges = []
        if index.db.execute(
            "SELECT 1 FROM sqlite_master WHERE name='merge_definitions'"
        ).fetchone():
            merges = index.db.execute(
                "SELECT r1,c1,r2,c2 FROM merge_definitions WHERE sheet=? LIMIT ?",
                (sheet, limits.header_regions + 1),
            ).fetchall()
            if len(merges) > limits.header_regions:
                self.coverage.update(
                    rowsScanned=len(self.rows),
                    capsHit=sorted({*caps, "header_regions"}),
                    complete=False,
                )
                return

        def candidate(row, cells):
            if not cells:
                return None
            headers = {c: literal(n) for c, n in cells.items()}
            if any(t is None or len(t) > 240 for t in headers.values()):
                return None
            columns = sorted(headers)
            if columns != list(range(columns[0], columns[-1] + 1)):
                return None
            words = [c for c, t in headers.items() if descriptive(t)]
            numeric_headers = [c for c in columns if c not in words]
            if numeric_headers:
                # A data row such as ['Region', '12', '17'] is insufficient.
                # Require two descriptive anchors and a literal consecutive
                # integer axis of at least three cells (e.g. ONS ages).
                if len(words) < 2 or len(numeric_headers) < 3:
                    return None
                try:
                    axis = [int(headers[c].strip()) for c in numeric_headers]
                except ValueError:
                    return None
                if numeric_headers != list(
                    range(numeric_headers[0], numeric_headers[-1] + 1)
                ) or axis != list(range(axis[0], axis[0] + len(axis))):
                    return None
            if not words:
                return None
            below = self.rows.get(row + 1, {})
            data_columns = [
                c
                for c in columns
                if c in below and numeric(below[c]) and not below[c].get("formula")
            ]
            if not data_columns:
                return None
            # Repeated numbers below a title do not prove a field header.
            # Declared tables already use their authoritative column metadata.
            if len(columns) < 2:
                return None
            self.coverage["candidates"] += 1
            if any(
                r1 <= row <= r2 and c1 <= columns[-1] and c2 >= columns[0]
                for r1, c1, r2, c2 in merges
            ):
                self.coverage["ambiguousCandidates"] += 1
                return None
            headers = {c: t for c, t in headers.items() if c in below}
            if len(headers) < 2:
                return None
            row_columns = [c for c in words if c in headers and c < min(data_columns)]
            if len(row_columns) > limits.header_row_labels:
                caps.add("header_row_labels")
                row_columns = []  # Never choose an arbitrary subset of row meaning.
            return headers, row_columns

        candidates = {
            row: found
            for row, cells in self.rows.items()
            if (found := candidate(row, cells))
        }
        for row, (headers, row_columns) in candidates.items():
            if len(self.bands) >= limits.header_bands:
                caps.add("header_bands")
                break
            # Split columns independently on missing cells. A complete new
            # textual section/header terminates the whole previous band.
            active = dict(headers)
            end = row
            for next_row in range(row + 1, max(self.rows, default=row) + 1):
                values = self.rows.get(next_row)
                if not values or next_row in candidates:
                    break
                if not any(numeric(n) or n.get("formula") for n in values.values()):
                    break
                if any(
                    r1 <= next_row <= r2 and c1 <= max(headers) and c2 >= min(headers)
                    for r1, c1, r2, c2 in merges
                ):
                    break
                # Omitted headers cannot continue below a hole or a new
                # nonnumeric value in a previously numeric data column.
                for col in list(active):
                    n = values.get(col)
                    if n is None or n.get("value") is None and not n.get("formula"):
                        del active[col]
                if active != headers:
                    break  # Conservative rectangle; do not leap over ragged rows.
                end = next_row
            if end > row:
                self.bands.append(Band(row, row + 1, end, headers, row_columns))
        self.coverage.update(
            rowsScanned=len(self.rows),
            bands=len(self.bands),
            capsHit=sorted(caps),
            complete=not caps,
        )

    def labels(self, row, column):
        bands = [
            b
            for b in self.bands
            if b.first_row <= row <= b.last_row and column in b.headers
        ]
        if len(bands) != 1:
            return []
        band = bands[0]
        result = []

        def add(text, r, c, role):
            result.append(
                dict(
                    text=text,
                    sheet=self.sheet,
                    cell=a1(r, c),
                    method="header_band",
                    role=role,
                    truncated=False,
                    certainty="heuristic",
                    field=band.headers[c],
                    fieldCell=a1(band.header_row, c),
                    headerRow=band.header_row,
                    region=(
                        f"{a1(band.first_row, min(band.headers))}:"
                        f"{a1(band.last_row, max(band.headers))}"
                    ),
                )
            )

        add(band.headers[column], band.header_row, column, "column")
        for col in band.row_columns:
            text = literal(self.rows.get(row, {}).get(col, {}))
            if text is not None and col < column:
                add(
                    text,
                    row,
                    col,
                    "row" if col == band.row_columns[0] else "row_context",
                )
        return result

    def records(self):
        return [asdict(band) for band in self.bands]
