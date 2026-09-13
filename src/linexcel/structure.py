"""What the file says about itself, before anything analyses it.

Extracted mechanically from analyzer.py: sheet dimensions, defined names and
the lightweight pre-analysis (``inspect_workbook``) that reads only package
headers, never formulas or values.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from typing import Any

from openpyxl import load_workbook

from linexcel.external import read_external_links
from linexcel.loader import MAX_CELLS_PER_SHEET, MAX_DENSE_CELLS, declared_cells
from linexcel.tables import _collect_defined_names, _force_dimensions

MAX_NODES_PER_SHEET = 400

#: Seconds per megabyte of uncompressed sheet XML, for everything that reads
#: the file: parsing, sweeping, grouping. Measured across workbooks from a
#: thousand cells to two hundred thousand formulas.
SECONDS_PER_SHEET_MB = 0.5
#: Seconds of evaluation per formula cell. Measured near 15 µs on chained
#: formulas; doubled here because a real workbook mixes in cross-sheet
#: references and lookups, and an estimate reads better slightly high than
#: an order of magnitude low.
SECONDS_PER_FORMULA = 30e-6
#: Below this the estimate is noise and nobody was going to wait anyway.
WORTH_MENTIONING_SECONDS = 5.0

_SHEET_PART_RE = re.compile(r"xl/worksheets/sheet\d+\.xml")


@dataclass
class Structure:
    sheet_dims: dict[str, tuple[int, int]]
    defined_names: dict[str, list]


def read_structure(data: bytes) -> Structure:
    owb = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    try:
        sheet_dims: dict[str, tuple[int, int]] = {}
        for ws in owb.worksheets:
            max_row, max_col = ws.max_row, ws.max_column
            if not max_row or not max_col:
                max_row, max_col = _force_dimensions(ws)
            sheet_dims[ws.title] = (max_row or 1, max_col or 1)
        defined_names = _collect_defined_names(owb)
    finally:
        owb.close()
    return Structure(sheet_dims, defined_names)


def sheet_bytes(data: bytes) -> int:
    """Uncompressed weight of the sheet parts, without unpacking one.

    The zip central directory carries each entry's real size, so this costs a
    read of the index — microseconds on a file that takes minutes to analyse.
    That is the whole point: an estimate nobody waits for.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            return sum(
                entry.file_size
                for entry in zf.infolist()
                if _SHEET_PART_RE.fullmatch(entry.filename)
            )
    except Exception:
        return 0


def count_formulas(data: bytes) -> int:
    """Number of formula cells, counted in the sheet XML.

    Unlike :func:`sheet_bytes` this unpacks the sheet parts, so it is not the
    free read the zip index gives. ``<f`` opens a formula element and nothing
    else in sheet XML, and a shared-formula slave is a real formula cell, so
    the raw byte count is the count.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            return sum(
                zf.read(entry.filename).count(b"<f")
                for entry in zf.infolist()
                if _SHEET_PART_RE.fullmatch(entry.filename)
            )
    except Exception:
        return 0


def estimate_seconds(data: bytes) -> float:
    """How long analysing this file is likely to take, in seconds.

    Two terms: reading the file scales with the weight of the sheet parts,
    and evaluating it scales with the number of formulas. The count costs an
    unpack of the sheets, so it is only paid when the weight alone already
    says the run will be long — a small file gets the cheap floor, which is
    all the warning it needs.
    """
    seconds = sheet_bytes(data) / 1_048_576 * SECONDS_PER_SHEET_MB
    if seconds < WORTH_MENTIONING_SECONDS:
        return seconds
    return seconds + count_formulas(data) * SECONDS_PER_FORMULA


def inspect_workbook(
    data: bytes,
    *,
    max_cells_per_sheet: int | None = None,
    max_nodes_per_sheet: int | None = None,
    max_dense_cells: int | None = None,
) -> dict[str, Any]:
    """What the file says about itself, before anything analyses it.

    Everything here is read from the package headers — sheet dimensions and
    external link declarations — so it costs milliseconds on a file that would
    take minutes to analyse. That is the point: it answers "is this going to
    be long, and will anything be left out?" *before* someone commits to
    finding out the slow way.

    Declared sizes, not real ones. A sheet claiming 17 billion cells holds
    nothing of the sort, and saying so is exactly the warning worth having.
    """
    cells_limit = (
        MAX_CELLS_PER_SHEET
        if max_cells_per_sheet is None
        else max_cells_per_sheet
    )
    nodes_limit = (
        MAX_NODES_PER_SHEET
        if max_nodes_per_sheet is None
        else max_nodes_per_sheet
    )
    dense_limit = MAX_DENSE_CELLS if max_dense_cells is None else max_dense_cells
    owb = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    try:
        sheets = []
        for ws in owb.worksheets:
            max_row, max_col = ws.max_row, ws.max_column
            if not max_row or not max_col:
                max_row, max_col = _force_dimensions(ws)
            rows, cols = max_row or 1, max_col or 1
            sheets.append(
                {
                    "name": ws.title,
                    "rows": rows,
                    "cols": cols,
                    "cells": rows * cols,
                    "state": ws.sheet_state,
                    "truncated": rows * cols > cells_limit,
                }
            )
    finally:
        owb.close()
    books = read_external_links(data)
    weight = sheet_bytes(data)
    return {
        "bytes": len(data),
        "sheetBytes": weight,
        "estimatedSeconds": round(estimate_seconds(data), 1),
        "sheets": sheets,
        "declaredCells": sum(s["cells"] for s in sheets),
        "externalWorkbooks": [b.name for b in books.values()],
        "densePathRefused": declared_cells(data) > dense_limit,
        "ceilings": {
            "cellsPerSheet": cells_limit,
            "nodesPerSheet": nodes_limit,
            "denseCells": dense_limit,
        },
    }
