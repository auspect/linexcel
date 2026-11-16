"""Data regions: formula-free rectangles of literal cells (index schema 5).

A region does not decide where values are stored: every literal input cell
that can be rebuilt exactly is a grid cell, inside a region or not. A region
gives a large block of saved values one ``data`` node (bounds, header row,
column schema, fingerprint and a bounded preview), and links each range node
that reads it to it, with the exact sub-range read.

The rule (lot 2 ter, approved 15 November 2026):

1. An Excel table: its body (without header and totals rows), with the
   table's column names. A column holding a formula in the body rows (a
   calculated column) is left out. Each run of adjacent remaining columns of
   at least ``REGION_MIN_CELLS`` cells is an ``excel_table`` region.
2. Outside tables, a column is mixed when a formula (or a declared array)
   lies between its first and last literal cell; it stays per cell. A
   formula above or below a column's literals (a computed title, a total
   row) does not exclude it. Adjacent columns that are not mixed form a
   rectangle, cut at any formula row inside it, keeping its largest band.
   The header is the first row, among the band's first ``HEADER_SCAN_ROWS``
   rows holding values, whose values are all text and fill at least half
   the columns, followed by a row holding a value that is not text. Rows
   above it (title, notes) stay outside the region. The region ends by
   column ``GRID_COLUMNS``.

Every region holds at least ``REGION_MIN_CELLS`` saved values, counted, not
taken from its area: a column with two values, in rows 1 and 1,048,576, is
not a region.

A region never contains a formula: this is checked again on the stored
cells before a region is kept. When the rule is wrong the values stay
right: a missed region leaves its cells per cell; a region too wide or a
wrong header only makes the description (schema, row count, preview) less
exact, since a sub-range is always the exact intersection with what a
formula reads.
"""

from __future__ import annotations

import hashlib
import json
from bisect import bisect_left
from collections import Counter
from typing import Any

from linexcel.cell_grid import GRID_COLUMNS
from linexcel.refs import a1, quote_sheet

#: Fewest saved values a region holds.
REGION_MIN_CELLS = 10_000
HEADER_SCAN_ROWS = 20
PREVIEW_ROWS, PREVIEW_COLUMNS = 10, 12


def _reference(sheet, r1, c1, r2, c2):
    return f"{quote_sheet(sheet)}!{a1(r1, c1)}:{a1(r2, c2)}"


class _Formulas:
    """Formula cells and declared arrays of one sheet, by column."""

    def __init__(self, positions, arrays):
        self.rows: dict[int, list[int]] = {}
        for row, col in positions:
            self.rows.setdefault(col, []).append(row)
        for rows in self.rows.values():
            rows.sort()
        self.arrays = arrays

    def in_rect(self, r1, c1, r2, c2) -> bool:
        for col in range(c1, c2 + 1):
            rows = self.rows.get(col)
            if rows:
                at = bisect_left(rows, r1)
                if at < len(rows) and rows[at] <= r2:
                    return True
        return any(
            not (b < r1 or a > r2 or d < c1 or c > c2) for a, c, b, d in self.arrays
        )

    def rows_in(self, r1, c1, r2, c2) -> set[int]:
        found = set()
        for col in range(c1, c2 + 1):
            rows = self.rows.get(col, [])
            at = bisect_left(rows, r1)
            while at < len(rows) and rows[at] <= r2:
                found.add(rows[at])
                at += 1
        for a, c, b, d in self.arrays:
            if not (b < r1 or a > r2 or d < c1 or c > c2):
                found.update(range(max(a, r1), min(b, r2) + 1))
        return found


def _runs(columns):
    run: list[int] = []
    for col in columns:
        if run and col != run[-1] + 1:
            yield run
            run = []
        run.append(col)
    if run:
        yield run


def _header(store, sheet, r1, c1, r2, c2):
    """``(header row, names)`` or ``(None, None)``; see the module rule."""
    rows: dict[int, dict[int, tuple]] = {}
    for row, col, value, kind in store.literal_values(
        sheet, (r1, c1, r2, c2), rows=HEADER_SCAN_ROWS + 1
    ):
        rows.setdefault(row, {})[col] = (value, kind)
    ordered = sorted(rows)
    width = c2 - c1 + 1
    for index, row in enumerate(ordered[:-1]):
        present = rows[row]
        if len(present) * 2 >= width and all(k == "text" for _, k in present.values()):
            following = rows[ordered[index + 1]]
            if any(k != "text" for _, k in following.values()):
                names = [
                    str(present[c][0]) if c in present else None
                    for c in range(c1, c2 + 1)
                ]
                return row, names
    return None, None


def detect_regions(store) -> list[dict]:
    """Candidate regions of every sheet, described, in sheet order."""
    tables: dict[str, list[dict]] = {}
    for (data,) in store.db.execute("SELECT data FROM table_definitions"):
        table = json.loads(data)
        tables.setdefault(table["sheet"], []).append(table)
    arrays: dict[str, list[tuple]] = {}
    for sheet, r1, c1, r2, c2 in store.db.execute("SELECT * FROM arrays"):
        arrays.setdefault(sheet, []).append((r1, c1, r2, c2))
    regions: list[dict[str, Any]] = []
    for sheet in store.import_sheets():
        spans = store.literal_spans(sheet)
        if not spans:
            continue
        formulas = _Formulas(store.formula_positions(sheet), arrays.get(sheet, []))
        sheet_tables = tables.get(sheet, [])
        for table in sheet_tables:
            regions.extend(_table_regions(sheet, table, formulas))
        mixed = {
            col
            for col, (first, last) in spans.items()
            if formulas.in_rect(first, col, last, col)
        }
        candidates = sorted(c for c in spans if c not in mixed and c <= GRID_COLUMNS)
        for run in _runs(candidates):
            c1, c2 = run[0], run[-1]
            r1 = min(spans[c][0] for c in run)
            r2 = max(spans[c][1] for c in run)
            if any(
                not (t["bounds"][2] < r1 or t["bounds"][0] > r2)
                and not (t["bounds"][3] < c1 or t["bounds"][1] > c2)
                for t in sheet_tables
            ):
                continue  # Excel tables follow their own rule
            cuts = sorted(formulas.rows_in(r1, c1, r2, c2))
            bands, start = [], r1
            for cut in cuts:
                if cut > start:
                    bands.append((start, cut - 1))
                start = cut + 1
            if start <= r2:
                bands.append((start, r2))
            if not bands:
                continue
            top, bottom = max(bands, key=lambda band: band[1] - band[0])
            header_row, headers = _header(store, sheet, top, c1, bottom, c2)
            body = header_row + 1 if header_row else top
            if (bottom - body + 1) * (c2 - c1 + 1) < REGION_MIN_CELLS:
                continue
            regions.append(
                {
                    "sheet": sheet,
                    "bounds": [body, c1, bottom, c2],
                    "regionKind": "formula_free_columns",
                    "headerRow": header_row,
                    "headers": headers,
                    "preambleRows": [top, header_row - 1]
                    if header_row and header_row > top
                    else None,
                }
            )
    kept = []
    for region in regions:
        r1, c1, r2, c2 = region["bounds"]
        sheet = region["sheet"]
        formulas = _Formulas(
            store.formula_positions(sheet, (r1, c1, r2, c2)), arrays.get(sheet, [])
        )
        if formulas.in_rect(r1, c1, r2, c2):
            continue  # never a formula inside a region
        if store.count_cells(sheet, (r1, c1, r2, c2)) < REGION_MIN_CELLS:
            continue  # an upper bound of its values: too sparse
        node = _describe(store, region)
        if node["cells"] >= REGION_MIN_CELLS:
            kept.append(node)
    return kept


def _table_regions(sheet, table, formulas):
    r1, c1, r2, c2 = table["bounds"]
    top, bottom = r1 + table["headerRows"], r2 - table["totalRows"]
    if top > bottom:
        return []
    excluded = [c for c in range(c1, c2 + 1) if formulas.in_rect(top, c, bottom, c)]
    regions = []
    columns = [c for c in range(c1, min(c2, GRID_COLUMNS) + 1) if c not in excluded]
    for run in _runs(columns):
        if (bottom - top + 1) * len(run) < REGION_MIN_CELLS:
            continue
        regions.append(
            {
                "sheet": sheet,
                "bounds": [top, run[0], bottom, run[-1]],
                "regionKind": "excel_table",
                "table": table["name"],
                "headerRow": r1 if table["headerRows"] else None,
                "headers": [table["headers"][c - c1] for c in run],
                "excludedFormulaColumns": [table["headers"][c - c1] for c in excluded],
            }
        )
    return regions


def _describe(store, region) -> dict:
    sheet = region["sheet"]
    r1, c1, r2, c2 = region["bounds"]
    digest = hashlib.sha256()
    types = {c: Counter() for c in range(c1, c2 + 1)}
    preview: dict[int, list] = {}
    cells = 0
    for row, col, value, kind in store.literal_values(sheet, (r1, c1, r2, c2)):
        cells += 1
        types[col][kind] += 1
        digest.update(
            json.dumps(
                [row, col, kind, value], ensure_ascii=False, default=str
            ).encode()
        )
        if col < c1 + PREVIEW_COLUMNS and (
            row in preview or len(preview) < PREVIEW_ROWS
        ):
            preview.setdefault(row, [None] * min(PREVIEW_COLUMNS, c2 - c1 + 1))[
                col - c1
            ] = value
    rows = r2 - r1 + 1
    schema = []
    headers = region.get("headers") or [None] * (c2 - c1 + 1)
    for offset, col in enumerate(range(c1, c2 + 1)):
        observed = dict(types[col])
        kinds = set(observed)
        schema.append(
            {
                "column": col,
                "name": headers[offset],
                "type": "empty"
                if not kinds
                else next(iter(kinds))
                if len(kinds) == 1
                else "mixed",
                "observed": observed,
                "empty": rows - sum(observed.values()),
            }
        )
    reference = _reference(sheet, r1, c1, r2, c2)
    node = {
        "id": "data:" + reference,
        "kind": "data",
        "sheet": sheet,
        "reference": reference,
        "bounds": [r1, c1, r2, c2],
        **{k: v for k, v in region.items() if k not in {"sheet", "bounds"}},
        "rowCount": rows,
        "columnCount": c2 - c1 + 1,
        "cells": cells,
        "schema": schema,
        "calculable": not any(t.get("unreadable") for t in types.values()),
        "fingerprint": "sha256:" + digest.hexdigest(),
        "preview": {
            "rows": [[row, *values] for row, values in sorted(preview.items())],
            "bounded": f"{PREVIEW_ROWS} rows x {PREVIEW_COLUMNS} columns",
        },
        "provenance": "derived_region",
        "dependencies": [],
        "diagnostics": [],
    }
    if not node["calculable"]:
        node["diagnostics"].append(
            "Region holds an unreadable cell: a calculation reading it is refused."
        )
    return node


def index_regions(store) -> dict:
    """Store each region as a ``data`` node; link the ranges that read it."""
    regions = detect_regions(store)
    by_sheet: dict[str, list[dict]] = {}
    for node in regions:
        store.put(node)
        by_sheet.setdefault(node["sheet"], []).append(node)
    links = 0
    if by_sheet:
        ranges = [
            node
            for _, node in store.range_nodes()
            if node.get("sheet") in by_sheet
            and not node.get("external")
            and node.get("bounds")
        ]
        for node in ranges:
            a, b, c, d = node["bounds"]
            found = []
            for region in by_sheet[node["sheet"]]:
                r1, c1, r2, c2 = region["bounds"]
                sub = (max(a, r1), max(b, c1), min(c, r2), min(d, c2))
                if sub[0] > sub[2] or sub[1] > sub[3]:
                    continue
                found.append(
                    {
                        "node": region["id"],
                        "subRange": _reference(node["sheet"], *sub),
                        "bounds": list(sub),
                        "coversWholeRange": sub == (a, b, c, d),
                    }
                )
                store.db.execute(
                    "INSERT OR IGNORE INTO edges VALUES (?,?)",
                    (region["id"], node["id"]),
                )
            if found:
                node["dataRegions"] = found
                store.put(node)
                links += len(found)
    return {
        "regions": len(regions),
        "rangeLinks": links,
        "rule": "lot2ter-v1",
        "minCells": REGION_MIN_CELLS,
    }
