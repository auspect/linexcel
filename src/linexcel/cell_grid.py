"""Per-sheet grids for literal input cells (index schema 5).

A literal cell used to be one compressed JSON record in the index's ``nodes``
table, with its own row in three SQLite indexes: several gigabytes for a large
data sheet. Schema 5 keeps such a cell as one value in a row of its sheet's
grid table, and rebuilds the node when a reader asks for it.

The grid accepts a cell only when the node rebuilt from it equals the node
the reader produced, field by field, except ``sourceValueProvenance``'s
``rawXmlValue`` (the XML spelling of the value, which the grid does not
keep). Anything else stays a node: formulas, array members, cells with a
diagnostic, a missing or unreadable value, an ISO date (``t="d"``), a cell
out of row order, and any field that does not have its usual form.

Layout, written in one forward pass by ``GridWriter``:

``grid_sheets(no, name, grid, width, cells, data)``
    one row per sheet with grid cells; ``data`` holds the sheet's provenance
    (workbook SHA-256, worksheet part, sheet state) and its hidden columns.
``g_NNN(row INTEGER PRIMARY KEY, h, c1 ... cW)``
    a sheet's values, SQLite-typed: numbers as saved (date cells keep their
    serial), text, errors as their text, booleans as 0 or 1. ``h`` is 1 for
    a hidden row. Columns beyond ``GRID_COLUMNS`` go to ``grid_wide``.
``grid_formats(id, storage_type, number_format, date_kind, address)``
    what a value needs to become a node again: its XML storage type, number
    format, whether it was converted to a date or a duration, and whether
    the XML gave the cell address.
``grid_runs(sheet_no, col, r1, r2, format)``
    the format of each column's grid cells, by runs of rows.
"""

from __future__ import annotations

import heapq
import json
import math
import re
from bisect import bisect_right
from collections.abc import Iterator
from typing import Any

from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
    from_excel,
)

from linexcel.refs import a1, parse_ref, quote_sheet

#: Wider sheets keep their extra columns in ``grid_wide``.
GRID_COLUMNS = 512
_ROWS_PER_WRITE = 2000
_TYPES = frozenset({"n", "s", "str", "inlineStr", "b", "e"})
_TEXT_TYPES = frozenset({"s", "str", "inlineStr", "e"})
_NODE_FIELDS = frozenset(
    {
        "id",
        "kind",
        "sheet",
        "cell",
        "row",
        "column",
        "formula",
        "sourceValueProvenance",
        "numberFormat",
        "sourceDataType",
        "context",
        "value",
        "cached_value",
        "cachedValueKind",
        "cached_comparison_value",
        "cache_status",
        "provenance",
        "dependencies",
        "diagnostics",
    }
)
_PROVENANCE_FIELDS = frozenset(
    {
        "workbookSha256",
        "worksheetPart",
        "sheet",
        "cell",
        "xmlCellAddress",
        "cellAddressOrigin",
        "role",
        "storageType",
        "valueElementPresent",
        "rawXmlValue",
        "status",
    }
)
_CONTEXT_FIELDS = frozenset({"sheetState", "rowHidden", "columnHidden"})
_NOT_DATE, _DATE, _DURATION = 0, 1, 2
_INT_LIMIT = 1 << 63

GRID_SCHEMA = """
    CREATE TABLE grid_sheets(no INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,
        grid TEXT NOT NULL,width INTEGER NOT NULL,cells INTEGER NOT NULL,
        data TEXT NOT NULL);
    CREATE TABLE grid_wide(sheet_no INTEGER NOT NULL,row INTEGER NOT NULL,
        col INTEGER NOT NULL,value,PRIMARY KEY(sheet_no,row,col)) WITHOUT ROWID;
    CREATE TABLE grid_formats(id INTEGER PRIMARY KEY,storage_type TEXT NOT NULL,
        number_format TEXT NOT NULL,date_kind INTEGER NOT NULL,
        address INTEGER NOT NULL);
    CREATE TABLE grid_runs(sheet_no INTEGER NOT NULL,col INTEGER NOT NULL,
        r1 INTEGER NOT NULL,r2 INTEGER NOT NULL,format INTEGER NOT NULL,
        PRIMARY KEY(sheet_no,col,r1)) WITHOUT ROWID;
"""


def cell_id(sheet: str, row: int, col: int) -> str:
    """The node id of a cell, as ``lazy._id`` spells it."""
    return f"{quote_sheet(sheet)}!{a1(row, col)}"


def _json_value(value):
    # lazy._json, for the date and duration values from_excel returns.
    from linexcel.lazy import _json

    return _json(value)


def _date_text(serial, kind, epoch):
    return _json_value(from_excel(serial, epoch, timedelta=kind == _DURATION))


class _Sheet:
    """Write state of one sheet's grid."""

    __slots__ = (
        "no",
        "name",
        "table",
        "width",
        "cells",
        "provenance",
        "hidden_columns",
        "last",
        "row",
        "row_hidden",
        "values",
        "pending",
        "wide",
        "runs",
        "closed_runs",
        "arrays",
        "active",
        "frozen",
        "created",
    )

    def __init__(self, no: int, name: str):
        self.no, self.name = no, name
        self.table = f"g_{no:03d}"
        self.width = 0
        self.cells = 0
        self.provenance: tuple | None = None
        self.hidden_columns: dict[int, bool] = {}
        self.last: tuple[int, int] | None = None
        self.row: int | None = None
        self.row_hidden = False
        self.values: dict[int, Any] = {}
        self.pending: list[tuple] = []
        self.wide: list[tuple] = []
        self.runs: dict[int, list[int]] = {}
        self.closed_runs: list[tuple] = []
        self.arrays: list[tuple[int, int, int, int]] = []  # heap, by first row
        self.active: list[tuple[int, int, int, int]] = []
        self.frozen = False
        self.created = False


class GridWriter:
    """Receives the reader's cell nodes; keeps reconstructible literals.

    ``offer`` is called for every cell node in reading order. It returns
    ``True`` when the cell went to its sheet's grid; the caller stores any
    other node in ``nodes``. ``demote`` is how the writer hands a grid cell
    back as a node when a later cell claims its position.
    """

    def __init__(self, db, *, epoch1904: bool, demote):
        self.db = db
        self.epoch1904 = epoch1904
        self.epoch = CALENDAR_MAC_1904 if epoch1904 else CALENDAR_WINDOWS_1900
        self.demote_node = demote
        self.formats: dict[tuple, int] = {}
        self.sheets: dict[str, _Sheet] = {}
        self.current: _Sheet | None = None
        self.closed: set[str] = set()
        #: Sheets in the order their first cell was read.
        self.order: list[str] = []

    # Sheets ---------------------------------------------------------------

    def _sheet(self, name: str) -> _Sheet:
        current = self.current
        if current is not None and current.name == name:
            return current
        if current is not None:
            self._close(current)
        sheet = self.sheets.get(name)
        if sheet is None:
            sheet = self.sheets[name] = _Sheet(len(self.sheets) + 1, name)
            self.order.append(name)
        if name in self.closed:
            # A sheet seen again later is out of reading order: keep its new
            # cells as nodes.
            sheet.frozen = True
        self.current = sheet
        return sheet

    def _close(self, sheet: _Sheet) -> None:
        self._flush_row(sheet)
        self._flush(sheet)
        for col, (r1, r2, fmt) in sheet.runs.items():
            sheet.closed_runs.append((sheet.no, col, r1, r2, fmt))
        sheet.runs = {}
        self._flush_runs(sheet, force=True)
        self.closed.add(sheet.name)

    def note_array(self, sheet_name: str, r1: int, c1: int, r2: int, c2: int):
        """A declared array or data table: its cells are never grid cells."""
        sheet = self._sheet(sheet_name)
        rect = (r1, c1, r2, c2)
        if sheet.row is not None and r1 <= sheet.row:
            if sheet.row <= r2:
                sheet.active.append(rect)
        else:
            heapq.heappush(sheet.arrays, rect)
        # Cells already read inside it (only possible out of reading order).
        if sheet.cells and sheet.last is not None and (r1, c1) <= sheet.last:
            for row, col in self._stored_positions(sheet, rect):
                self._demote(sheet, row, col)

    # Cells ----------------------------------------------------------------

    def offer(self, node: dict) -> bool:
        name, row, col = node.get("sheet"), node.get("row"), node.get("column")
        if (
            name is None
            or type(row) is not int
            or type(col) is not int
            or node.get("kind") not in {"cell", "input", "array_result"}
        ):
            return False
        sheet = self._sheet(name)
        position = (row, col)
        if sheet.last is not None and position <= sheet.last:
            # Out of reading order, or a second cell at one position: the
            # node keeps it, and a grid cell there becomes a node first so
            # the later cell replaces it as it would replace a node.
            self._demote(sheet, row, col)
            return False
        sheet.last = position
        if row != sheet.row:
            self._start_row(sheet, row)
        if sheet.frozen:
            return False
        entry = self._entry(sheet, node)
        if entry is None:
            return False
        value, fmt, hidden = entry
        if sheet.values and hidden != sheet.row_hidden:
            return False
        sheet.row_hidden = hidden
        if not sheet.created:
            self.db.execute(f'CREATE TABLE "{sheet.table}" (row INTEGER PRIMARY KEY,h)')
            sheet.created = True
        if col > GRID_COLUMNS:
            sheet.wide.append((sheet.no, row, col, value))
        else:
            if col > sheet.width:
                self._widen(sheet, col)
            sheet.values[col] = value
        sheet.values.setdefault(0, None)  # the row exists even with wide cells only
        run = sheet.runs.get(col)
        if run is not None and run[2] == fmt:
            run[1] = row
        else:
            if run is not None:
                sheet.closed_runs.append((sheet.no, col, run[0], run[1], run[2]))
                self._flush_runs(sheet)
            sheet.runs[col] = [row, row, fmt]
        sheet.cells += 1
        return True

    def _start_row(self, sheet: _Sheet, row: int) -> None:
        self._flush_row(sheet)
        sheet.row, sheet.row_hidden = row, False
        if sheet.active:
            sheet.active = [a for a in sheet.active if a[2] >= row]
        while sheet.arrays and sheet.arrays[0][0] <= row:
            rect = heapq.heappop(sheet.arrays)
            if rect[2] >= row:
                sheet.active.append(rect)

    def _entry(self, sheet: _Sheet, node: dict):
        """``(stored value, format id, row hidden)``, or ``None``."""
        if node.keys() != _NODE_FIELDS:
            return None
        if (
            node["kind"] != "input"
            or node["formula"] is not None
            or node["provenance"] != "saved_input"
            or node["cache_status"] != "saved"
            or node["dependencies"] != []
            or node["diagnostics"] != []
        ):
            return None
        kind = node["sourceDataType"]
        if kind not in _TYPES or not isinstance(node["numberFormat"], str):
            return None
        if node["cachedValueKind"] != ("error" if kind == "e" else "scalar"):
            return None
        row, col, name = node["row"], node["column"], sheet.name
        cell = node["cell"]
        if cell != a1(row, col) or node["id"] != f"{quote_sheet(name)}!{cell}":
            return None
        for top, left, bottom, right in sheet.active:
            if left <= col <= right and top <= row <= bottom:
                return None
        provenance = node["sourceValueProvenance"]
        if (
            type(provenance) is not dict
            or provenance.keys() != _PROVENANCE_FIELDS
            or provenance["role"] != "input"
            or provenance["storageType"] != kind
            or provenance["sheet"] != name
            or provenance["cell"] != cell
            or provenance["status"] != "present"
            or provenance["valueElementPresent"] is not (kind != "inlineStr")
        ):
            return None
        if (
            provenance["xmlCellAddress"] == cell
            and provenance["cellAddressOrigin"] == "stored_xml"
        ):
            address = 1
        elif (
            provenance["xmlCellAddress"] is None
            and provenance["cellAddressOrigin"] == "inferred"
        ):
            address = 0
        else:
            return None
        context = node["context"]
        if type(context) is not dict or context.keys() != _CONTEXT_FIELDS:
            return None
        hidden, column_hidden = context["rowHidden"], context["columnHidden"]
        if type(hidden) is not bool or type(column_hidden) is not bool:
            return None
        sheet_key = (
            provenance["workbookSha256"],
            provenance["worksheetPart"],
            context["sheetState"],
        )
        if sheet.provenance is None:
            if not all(isinstance(v, str) for v in sheet_key):
                return None
            sheet.provenance = sheet_key
        elif sheet_key != sheet.provenance:
            return None
        if sheet.hidden_columns.setdefault(col, column_hidden) is not column_hidden:
            return None
        stored, date_kind = self._value(kind, node)
        if date_kind is None:
            return None
        key = (kind, node["numberFormat"], date_kind, address)
        fmt = self.formats.get(key)
        if fmt is None:
            fmt = self.formats[key] = len(self.formats) + 1
        return stored, fmt, hidden

    def _value(self, kind: str, node: dict):
        """``(stored value, date kind)``; ``(None, None)`` when not exact."""
        value = node["value"]
        cached = node["cached_value"]
        if type(cached) is not type(value) or cached != value:
            return None, None
        comparison = node["cached_comparison_value"]
        if kind in _TEXT_TYPES:
            if type(value) is not str or type(comparison) is not str:
                return None, None
            if comparison != value:
                return None, None
            if not value.isascii():
                try:
                    value.encode("utf-8")
                except UnicodeEncodeError:
                    return None, None
            return value, _NOT_DATE
        if kind == "b":
            if type(value) is not bool or comparison is not value:
                return None, None
            return int(value), _NOT_DATE
        # kind == "n": the saved number is the comparison value.
        if type(comparison) is int:
            if not -_INT_LIMIT <= comparison < _INT_LIMIT:
                return None, None
        elif type(comparison) is float:
            if not math.isfinite(comparison) or (
                comparison == 0 and math.copysign(1.0, comparison) < 0
            ):
                return None, None
        else:
            return None, None
        if type(value) is type(comparison) and value == comparison:
            return comparison, _NOT_DATE
        if type(value) is not str:
            return None, None
        for date_kind in (_DATE, _DURATION):
            try:
                if _date_text(comparison, date_kind, self.epoch) == value:
                    return comparison, date_kind
            except (ValueError, OverflowError, TypeError):
                return None, None
        return None, None

    # Writes ---------------------------------------------------------------

    def _widen(self, sheet: _Sheet, col: int) -> None:
        # Queued rows were built at the old width; the open row is still in
        # ``values`` and is written later at the new width.
        self._flush(sheet)
        for c in range(sheet.width + 1, col + 1):
            self.db.execute(f'ALTER TABLE "{sheet.table}" ADD COLUMN c{c}')
        sheet.width = col

    def _flush_row(self, sheet: _Sheet) -> None:
        if not sheet.values:
            return
        values = sheet.values
        sheet.pending.append(
            (
                sheet.row,
                1 if sheet.row_hidden else None,
                *(values.get(c) for c in range(1, sheet.width + 1)),
            )
        )
        sheet.values = {}
        if len(sheet.pending) >= _ROWS_PER_WRITE:
            self._flush(sheet)

    def _flush(self, sheet: _Sheet) -> None:
        if sheet.pending:
            columns = "".join(f",c{c}" for c in range(1, sheet.width + 1))
            marks = ",?" * sheet.width
            self.db.executemany(
                f'INSERT INTO "{sheet.table}" (row,h{columns}) VALUES (?,?{marks})',
                sheet.pending,
            )
            sheet.pending = []
        if sheet.wide:
            self.db.executemany("INSERT INTO grid_wide VALUES (?,?,?,?)", sheet.wide)
            sheet.wide = []

    def _flush_runs(self, sheet: _Sheet, *, force=False) -> None:
        if sheet.closed_runs and (force or len(sheet.closed_runs) >= _ROWS_PER_WRITE):
            self.db.executemany(
                "INSERT INTO grid_runs VALUES (?,?,?,?,?)", sheet.closed_runs
            )
            sheet.closed_runs = []

    # Demotion -------------------------------------------------------------

    def _stored_positions(self, sheet: _Sheet, rect) -> list[tuple[int, int]]:
        r1, c1, r2, c2 = rect
        found = [
            (sheet.row, c)
            for c in sheet.values
            if c and sheet.row is not None and r1 <= sheet.row <= r2 and c1 <= c <= c2
        ]
        if sheet.created:
            self._flush(sheet)
            low, high = c1, min(c2, sheet.width)
            if low <= high:
                columns = ",".join(f"c{c}" for c in range(low, high + 1))
                for row, *values in self.db.execute(
                    f'SELECT row,{columns} FROM "{sheet.table}" '
                    "WHERE row BETWEEN ? AND ?",
                    (r1, r2),
                ):
                    found.extend(
                        (row, low + i) for i, v in enumerate(values) if v is not None
                    )
            found.extend(
                self.db.execute(
                    "SELECT row,col FROM grid_wide WHERE sheet_no=? "
                    "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ?",
                    (sheet.no, r1, r2, c1, c2),
                ).fetchall()
            )
        return sorted(set(found))

    def _demote(self, sheet: _Sheet, row: int, col: int) -> None:
        """Move one grid cell, if any, to the ``nodes`` table."""
        if not sheet.cells:
            return
        value, fmt, hidden = None, None, None
        if sheet.row == row and col in sheet.values and col:
            value, hidden = sheet.values.pop(col), sheet.row_hidden
        elif sheet.row == row and col > GRID_COLUMNS:
            for i, item in enumerate(sheet.wide):
                if item[1] == row and item[2] == col:
                    value, hidden = item[3], sheet.row_hidden
                    del sheet.wide[i]
                    break
        if hidden is None and sheet.created:
            self._flush(sheet)
            if col > GRID_COLUMNS:
                hit = self.db.execute(
                    "SELECT value FROM grid_wide WHERE sheet_no=? AND row=? AND col=?",
                    (sheet.no, row, col),
                ).fetchone()
                if hit is not None and hit[0] is not None:
                    self.db.execute(
                        "DELETE FROM grid_wide WHERE sheet_no=? AND row=? AND col=?",
                        (sheet.no, row, col),
                    )
                    value = hit[0]
                    hidden = self._row_hidden(sheet, row)
            elif col <= sheet.width:
                hit = self.db.execute(
                    f'SELECT h,c{col} FROM "{sheet.table}" WHERE row=?', (row,)
                ).fetchone()
                if hit is not None and hit[1] is not None:
                    self.db.execute(
                        f'UPDATE "{sheet.table}" SET c{col}=NULL WHERE row=?', (row,)
                    )
                    value, hidden = hit[1], bool(hit[0])
        if hidden is None:
            return
        fmt = self._format_at(sheet, row, col)
        sheet.cells -= 1
        formats = {v: k for k, v in self.formats.items()}
        provenance = sheet.provenance
        assert provenance is not None and fmt is not None
        node = synthesize(
            sheet.name,
            row,
            col,
            value,
            formats[fmt],
            hidden,
            sheet.hidden_columns.get(col, False),
            provenance,
            self.epoch,
        )
        self.demote_node(node)

    def _row_hidden(self, sheet: _Sheet, row: int) -> bool:
        hit = self.db.execute(
            f'SELECT h FROM "{sheet.table}" WHERE row=?', (row,)
        ).fetchone()
        return bool(hit and hit[0])

    def _format_at(self, sheet: _Sheet, row: int, col: int) -> int | None:
        run = sheet.runs.get(col)
        if run is not None and run[0] <= row:
            return run[2]
        for _, c, r1, r2, fmt in reversed(sheet.closed_runs):
            if c == col and r1 <= row <= r2:
                return fmt
        self._flush_runs(sheet, force=True)
        hit = self.db.execute(
            "SELECT format FROM grid_runs WHERE sheet_no=? AND col=? AND r1<=? "
            "ORDER BY r1 DESC LIMIT 1",
            (sheet.no, col, row),
        ).fetchone()
        return hit[0] if hit else None

    def discard(self, name: str) -> None:
        """Drop one sheet's grid cells; later cells of that sheet stay nodes."""
        sheet = self.sheets.get(name)
        if sheet is None:
            return
        if self.current is sheet:
            self.current = None
        if sheet.created:
            self.db.execute(f'DROP TABLE "{sheet.table}"')
        self.db.execute("DELETE FROM grid_wide WHERE sheet_no=?", (sheet.no,))
        self.db.execute("DELETE FROM grid_runs WHERE sheet_no=?", (sheet.no,))
        sheet.created, sheet.cells, sheet.width = False, 0, 0
        sheet.values, sheet.pending, sheet.wide = {}, [], []
        sheet.runs, sheet.closed_runs = {}, []
        sheet.frozen = True
        self.closed.add(name)

    def close(self) -> None:
        """Write what is buffered and describe each grid; idempotent."""
        if self.current is not None:
            self._close(self.current)
            self.current = None
        self.db.executemany(
            "INSERT OR REPLACE INTO grid_formats VALUES (?,?,?,?,?)",
            [(fmt, *key) for key, fmt in self.formats.items()],
        )
        for sheet in self.sheets.values():
            if not sheet.created:
                continue
            sha, part, state = sheet.provenance or (None, None, None)
            self.db.execute(
                "INSERT OR REPLACE INTO grid_sheets VALUES (?,?,?,?,?,?)",
                (
                    sheet.no,
                    sheet.name,
                    sheet.table,
                    sheet.width,
                    sheet.cells,
                    json.dumps(
                        {
                            "workbookSha256": sha,
                            "worksheetPart": part,
                            "sheetState": state,
                            "hiddenColumns": sorted(
                                c for c, flag in sheet.hidden_columns.items() if flag
                            ),
                            "epoch1904": self.epoch1904,
                        }
                    ),
                ),
            )


def synthesize(
    sheet, row, col, value, fmt, hidden, column_hidden, provenance, epoch
) -> dict:
    """The node of one grid cell, as ``lazy_store._unpack`` returns it.

    Same fields, values and key order as the node the reader produced and
    the ``nodes`` table stored, except ``rawXmlValue``, which the grid does
    not keep: the key is left out rather than set to ``None``.
    """
    kind, number_format, date_kind, address = fmt
    cell = a1(row, col)
    if kind == "b":
        value = bool(value)
    comparison = value
    if date_kind:
        value = _date_text(value, date_kind, epoch)
    sha, part, state = provenance
    node = {
        "id": f"{quote_sheet(sheet)}!{cell}",
        "kind": "input",
        "sheet": sheet,
        "cell": cell,
        "row": row,
        "column": col,
        "formula": None,
        "sourceValueProvenance": {
            "workbookSha256": sha,
            "worksheetPart": part,
            "sheet": sheet,
            "cell": cell,
            "xmlCellAddress": cell if address else None,
            "cellAddressOrigin": "stored_xml" if address else "inferred",
            "role": "input",
            "storageType": kind,
            "valueElementPresent": kind != "inlineStr",
            "status": "present",
        },
        "numberFormat": number_format,
        "sourceDataType": kind,
        "context": {
            "sheetState": state,
            "rowHidden": hidden,
            "columnHidden": column_hidden,
        },
        "value": value,
        "cachedValueKind": "error" if kind == "e" else "scalar",
    }
    if date_kind:
        node["cached_comparison_value"] = comparison
    node.update(
        cache_status="saved",
        provenance="saved_input",
        dependencies=[],
        diagnostics=[],
        cached_value=value,
    )
    if not date_kind:
        node["cached_comparison_value"] = value
    return node


class _ReadSheet:
    __slots__ = ("no", "name", "table", "width", "cells", "provenance", "hidden")

    def __init__(self, no, name, table, width, cells, data):
        self.no, self.name, self.table = no, name, table
        self.width, self.cells = width, cells
        data = json.loads(data)
        self.provenance = (
            data["workbookSha256"],
            data["worksheetPart"],
            data["sheetState"],
        )
        self.hidden = frozenset(data["hiddenColumns"])


class GridReader:
    """Reads grid cells back, as raw values or as nodes."""

    def __init__(self, db, epoch1904: bool):
        self.db = db
        self.epoch = CALENDAR_MAC_1904 if epoch1904 else CALENDAR_WINDOWS_1900
        self.sheets = {
            row[1]: _ReadSheet(*row)
            for row in db.execute(
                "SELECT no,name,grid,width,cells,data FROM grid_sheets ORDER BY no"
            )
        }
        self.formats = {
            row[0]: tuple(row[1:]) for row in db.execute("SELECT * FROM grid_formats")
        }

    @property
    def total(self) -> int:
        return sum(sheet.cells for sheet in self.sheets.values())

    def locate(self, node_id) -> tuple[str, int, int] | None:
        """``(sheet, row, col)`` when ``node_id`` names a cell of a grid sheet."""
        if not isinstance(node_id, str) or "!" not in node_id or not self.sheets:
            return None
        rect = parse_ref(node_id)
        if (
            rect is None
            or rect.ncells != 1
            or rect.sheet not in self.sheets
            or cell_id(rect.sheet, rect.r1, rect.c1) != node_id
        ):
            return None
        return rect.sheet, rect.r1, rect.c1

    def node(self, sheet: str, row: int, col: int, value, fmt, hidden) -> dict:
        info = self.sheets[sheet]
        return synthesize(
            sheet,
            row,
            col,
            value,
            self.formats[fmt],
            hidden,
            col in info.hidden,
            info.provenance,
            self.epoch,
        )

    def search_text(self, raw) -> str:
        """What a node's ``search_value`` holds: the text of its value."""
        value, fmt, _ = raw
        kind, _, date_kind, _ = self.formats[fmt]
        if kind == "b":
            return str(bool(value))
        if date_kind:
            return _date_text(value, date_kind, self.epoch)
        return str(value)

    def get(self, sheet: str, row: int, col: int) -> dict | None:
        """The node of one grid cell, or ``None``."""
        raw = self.raw(sheet, row, col)
        return None if raw is None else self.node(sheet, row, col, *raw)

    def raw(self, sheet: str, row: int, col: int):
        """``(value, format id, row hidden)`` of one grid cell, or ``None``."""
        info = self.sheets.get(sheet)
        if info is None:
            return None
        if col > GRID_COLUMNS:
            hit = self.db.execute(
                "SELECT value FROM grid_wide WHERE sheet_no=? AND row=? AND col=?",
                (info.no, row, col),
            ).fetchone()
            if hit is None or hit[0] is None:
                return None
            hidden = self.db.execute(
                f'SELECT h FROM "{info.table}" WHERE row=?', (row,)
            ).fetchone()
            value, flag = hit[0], hidden[0] if hidden else None
        else:
            if col > info.width:
                return None
            hit = self.db.execute(
                f'SELECT h,c{col} FROM "{info.table}" WHERE row=?', (row,)
            ).fetchone()
            if hit is None or hit[1] is None:
                return None
            flag, value = hit
        fmt = self.db.execute(
            "SELECT format FROM grid_runs WHERE sheet_no=? AND col=? AND r1<=? "
            "ORDER BY r1 DESC LIMIT 1",
            (info.no, col, row),
        ).fetchone()[0]
        return value, fmt, bool(flag)

    def _formats(self, info, r1, r2, columns):
        """Per column, ``(row starts, formats)`` of the runs over rows r1..r2."""
        runs: dict[int, tuple[list[int], list[int]]] = {}
        for col in columns:
            first = self.db.execute(
                "SELECT r1,format FROM grid_runs WHERE sheet_no=? AND col=? "
                "AND r1<=? ORDER BY r1 DESC LIMIT 1",
                (info.no, col, r1),
            ).fetchone()
            starts, formats = ([first[0]], [first[1]]) if first else ([], [])
            for start, fmt in self.db.execute(
                "SELECT r1,format FROM grid_runs WHERE sheet_no=? AND col=? "
                "AND r1>? AND r1<=? ORDER BY r1",
                (info.no, col, r1, r2),
            ):
                starts.append(start)
                formats.append(fmt)
            if starts:
                runs[col] = (starts, formats)
        return runs

    def columns_with_runs(self, info, c1, c2) -> list[int]:
        return [
            row[0]
            for row in self.db.execute(
                "SELECT DISTINCT col FROM grid_runs WHERE sheet_no=? "
                "AND col BETWEEN ? AND ? ORDER BY col",
                (info.no, c1, c2),
            )
        ]

    def scan(self, sheet: str, bounds=None) -> Iterator[tuple]:
        """``(row, col, value, format id, row hidden)``, row by row.

        ``bounds`` is ``(r1, c1, r2, c2)``; ``None`` reads the whole sheet.
        """
        info = self.sheets.get(sheet)
        if info is None or not info.cells:
            return
        r1, c1, r2, c2 = bounds or (1, 1, 1 << 31, 1 << 31)
        if r1 > r2 or c1 > c2:
            return
        columns = self.columns_with_runs(info, c1, c2)
        if not columns:
            return
        c1, c2 = columns[0], columns[-1]
        runs = self._formats(info, r1, r2, columns)
        narrow = (c1, min(c2, info.width)) if c1 <= info.width else None
        wide = c2 > GRID_COLUMNS
        if narrow:
            low, high = narrow
            names = ",".join(f"c{c}" for c in range(low, high + 1))
            rows = self.db.execute(
                f'SELECT row,h,{names} FROM "{info.table}" '
                "WHERE row BETWEEN ? AND ? ORDER BY row",
                (r1, r2),
            )
        else:
            low, rows = c1, iter(())
        extra = (
            self.db.execute(
                "SELECT row,col,value FROM grid_wide WHERE sheet_no=? "
                "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ? ORDER BY row,col",
                (info.no, r1, r2, max(c1, GRID_COLUMNS + 1), c2),
            )
            if wide
            else iter(())
        )

        def fmt_of(row, col):
            starts, formats = runs[col]
            return formats[bisect_right(starts, row) - 1]

        def narrow_cells():
            for row, flag, *values in rows:
                hidden = bool(flag)
                for i, value in enumerate(values):
                    if value is not None:
                        col = low + i
                        yield row, col, value, fmt_of(row, col), hidden

        def wide_cells():
            for row, col, value in extra:
                if value is None:
                    continue
                flag = self.db.execute(
                    f'SELECT h FROM "{info.table}" WHERE row=?', (row,)
                ).fetchone()
                yield row, col, value, fmt_of(row, col), bool(flag and flag[0])

        if wide and narrow:
            yield from heapq.merge(narrow_cells(), wide_cells(), key=_position)
        elif wide:
            yield from wide_cells()
        else:
            yield from narrow_cells()

    def count(self, sheet: str, bounds=None) -> int:
        info = self.sheets.get(sheet)
        if info is None or not info.cells:
            return 0
        if bounds is None:
            return info.cells
        r1, c1, r2, c2 = bounds
        total = 0
        low, high = c1, min(c2, info.width)
        if low <= high:
            for start in range(low, high + 1, 500):
                stop = min(high, start + 499)
                total += (
                    self.db.execute(
                        "SELECT "
                        + "+".join(f"count(c{c})" for c in range(start, stop + 1))
                        + f' FROM "{info.table}" WHERE row BETWEEN ? AND ?',
                        (r1, r2),
                    ).fetchone()[0]
                    or 0
                )
        if c2 > GRID_COLUMNS:
            total += self.db.execute(
                "SELECT count(*) FROM grid_wide WHERE sheet_no=? AND row BETWEEN ? "
                "AND ? AND col BETWEEN ? AND ? AND value IS NOT NULL",
                (info.no, r1, r2, max(c1, GRID_COLUMNS + 1), c2),
            ).fetchone()[0]
        return total


def _position(item):
    return item[0], item[1]


_SQLITE_REAL = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")


def sqlite_real(text: str) -> float:
    """``CAST(text AS REAL)`` for the numeric spellings SQLite accepts."""
    match = _SQLITE_REAL.match(text.strip())
    if not match:
        return 0.0
    try:
        return float(match.group())
    except (ValueError, OverflowError):
        return 0.0
