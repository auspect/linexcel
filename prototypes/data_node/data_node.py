"""Data node prototype (batch 2): one node per formula-free region.

Isolated from the product: nothing in ``src/linexcel`` imports this module and
it is not packaged. It answers four questions with measurements:

1. Can serve's streaming reader feed a compact value store instead of one
   zlib/JSON blob per cell (F06)?  ``build_store`` reuses
   ``lazy_store._read_source`` unchanged and only replaces its storage sink.
2. How is a data region recognised (Excel table, formula-free columns,
   header row), and what happens to mixed cases?  ``detect_regions``.
3. Can formulas point to the region with the sub-range they read?  ``link``.
4. Can targeted calculation load only the sub-range read, in batches, while
   keeping linexcel's closure, refusal and comparison logic?  ``evaluate``.

Store layout (one SQLite file):
  sheets(name, grid, width, literal_cells, formula_cells)
  g_<n>(row INTEGER PRIMARY KEY, c1..cW)  literal values, SQLite-typed:
      INTEGER/REAL numbers (date cells keep their serial), TEXT strings.
  wide(sheet, row, col, value)           literal cells beyond column 512
  exceptions(sheet, row, col, kind, value_json)
      bool | error | iso_date | format | unreadable | rich value kinds; an
      ``unreadable`` cell makes any calculation that reads it refuse.
  columns(sheet, col, stats_json)        per-column counts, types, formats
  regions(id, sheet, r1, c1, r2, c2, data_json)

Values are the saved file values: nothing here recalculates. A cell whose
value is unreadable stays unknown (exception ``unreadable``), never blank.
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
import time
from collections import Counter
from pathlib import Path
from typing import Any

GRID_COLUMNS = 512  # wider sheets keep their extra columns in ``wide``
REGION_MIN_CELLS = 10_000  # prototype threshold: smaller blocks stay per cell
CHUNK_ROWS = 10_000  # engine batch size (formualizer 0.14.0 is superlinear per call)
HEADER_SCAN_ROWS = 20


def _dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


MAX_FORMAT_RUNS = 1000  # per column; beyond, the column is marked fragmented


class _ColumnStats:
    __slots__ = (
        "literal",
        "formula",
        "first",
        "last",
        "types",
        "formats",
        "lo",
        "hi",
        "runs",
        "fragmented",
    )

    def __init__(self):
        self.literal = self.formula = 0
        self.first = self.last = None
        self.types: Counter = Counter()
        self.formats: Counter = Counter()
        self.lo = self.hi = None
        # Number format by runs of literal rows: [[first_row, last_row, format]].
        self.runs: list[list] = []
        self.fragmented = False

    def format_at(self, row: int, fmt: str):
        if self.runs and self.runs[-1][2] == fmt:
            self.runs[-1][1] = row
        elif len(self.runs) < MAX_FORMAT_RUNS:
            self.runs.append([row, row, fmt])
        else:
            self.fragmented = True

    def as_dict(self):
        return {
            "literalCells": self.literal,
            "formulaCells": self.formula,
            "firstRow": self.first,
            "lastRow": self.last,
            "types": dict(self.types),
            "formats": dict(self.formats.most_common(4)),
            "formatRuns": self.runs,
            "formatRunsFragmented": self.fragmented,
            "min": self.lo,
            "max": self.hi,
        }


class _Sink:
    """Receives serve's node dicts; writes literals to per-sheet grids."""

    def __init__(self, db: sqlite3.Connection):
        self.db = db
        self.sheet = None
        self.sheets: dict[str, dict] = {}
        self.columns: dict[str, dict[int, _ColumnStats]] = {}
        self.pending: list[tuple] = []
        self.exceptions: list[tuple] = []
        self.row = None
        self.values: dict[int, Any] = {}
        self.width = 0
        self.statement = ""

    def _start(self, sheet: str):
        self._flush_row()
        self._flush()
        number = len(self.sheets) + 1
        grid = f"g_{number:03d}"
        self.db.execute(f'CREATE TABLE "{grid}" (row INTEGER PRIMARY KEY)')
        self.sheets[sheet] = {
            "name": sheet,
            "grid": grid,
            "width": 0,
            "literal": 0,
            "formula": 0,
        }
        self.columns[sheet] = {}
        self.sheet, self.row, self.values, self.width = sheet, None, {}, 0
        self._statement()

    def _statement(self):
        grid = self.sheets[self.sheet]["grid"]
        columns = "".join(f",c{c}" for c in range(1, self.width + 1))
        marks = ",?" * self.width
        self.statement = f'INSERT INTO "{grid}" (row{columns}) VALUES (?{marks})'

    def _widen(self, column: int):
        # Complete rows were queued at the old width; the current row is still
        # open in self.values and is written later at the new width.
        self._flush()
        grid = self.sheets[self.sheet]["grid"]
        for c in range(self.width + 1, column + 1):
            self.db.execute(f'ALTER TABLE "{grid}" ADD COLUMN c{c}')
        self.width = self.sheets[self.sheet]["width"] = column
        self._statement()

    def _flush_row(self):
        if self.row is None or not self.values:
            return
        self.pending.append(
            (self.row, *(self.values.get(c) for c in range(1, self.width + 1)))
        )
        self.values = {}
        if len(self.pending) >= 2000:
            self._flush()

    def _flush(self):
        if self.pending:
            self.db.executemany(self.statement, self.pending)
            self.pending = []
        if self.exceptions:
            self.db.executemany(
                "INSERT INTO exceptions VALUES (?,?,?,?,?)", self.exceptions
            )
            self.exceptions = []

    def put(self, node: dict, *, only_if_absent: bool = False):
        if node.get("kind") not in {"cell", "input"} or not node.get("sheet"):
            return  # defined names stay ordinary nodes
        sheet, row, col = node["sheet"], node["row"], node["column"]
        if sheet != self.sheet:
            self._start(sheet)
        stats = self.columns[sheet].setdefault(col, _ColumnStats())
        stats.first = row if stats.first is None else min(stats.first, row)
        stats.last = row if stats.last is None else max(stats.last, row)
        info = self.sheets[sheet]
        if node["kind"] == "cell":  # a formula: stays a per-cell node
            stats.formula += 1
            info["formula"] += 1
            return
        info["literal"] += 1
        stats.literal += 1
        kind = node.get("sourceDataType")
        fmt = node.get("numberFormat") or "General"
        stats.formats[fmt] += 1
        if node.get("diagnostics"):
            self.exceptions.append(
                (sheet, row, col, "unreadable", _dump(node["diagnostics"]))
            )
            stats.types["unreadable"] += 1
            return
        value = node.get("cached_comparison_value", node.get("value"))
        stored = value
        if kind == "b":
            self.exceptions.append((sheet, row, col, "bool", _dump(value)))
            stored = None
        elif kind == "e":
            self.exceptions.append((sheet, row, col, "error", _dump(value)))
            stored = None
        elif kind == "d":
            self.exceptions.append(
                (sheet, row, col, "iso_date", _dump(node.get("value")))
            )
        elif isinstance(value, (dict, list)):
            self.exceptions.append((sheet, row, col, "json", _dump(value)))
            stored = None
        stats.types[kind or "n"] += 1
        if isinstance(stored, (int, float)) and not isinstance(stored, bool):
            stats.lo = stored if stats.lo is None else min(stats.lo, stored)
            stats.hi = stored if stats.hi is None else max(stats.hi, stored)
        stats.format_at(row, fmt)
        if stats.fragmented:
            self.exceptions.append((sheet, row, col, "format", _dump(fmt)))
        if col > GRID_COLUMNS:
            self.db.execute(
                "INSERT INTO wide VALUES (?,?,?,?)", (sheet, row, col, stored)
            )
            return
        if row != self.row:
            self._flush_row()
            self.row = row
        if col > self.width:
            self._widen(col)
            self.row = row
        self.values[col] = stored

    def close(self):
        self._flush_row()
        self._flush()
        for sheet, info in self.sheets.items():
            self.db.execute(
                "INSERT INTO sheets VALUES (?,?,?,?,?)",
                (sheet, info["grid"], info["width"], info["literal"], info["formula"]),
            )
            for col, stats in self.columns[sheet].items():
                self.db.execute(
                    "INSERT INTO columns VALUES (?,?,?)",
                    (sheet, col, _dump(stats.as_dict())),
                )
        self.db.commit()


class _CaptureStore:
    """The minimal store surface ``_read_source`` writes to."""

    def __init__(self, scratch: Path, sink: _Sink):
        from linexcel.lazy_store import _schema

        self.db = sqlite3.connect(scratch)
        _schema(self.db)
        # A private scratch file: hold its lock. On Windows every autocommit
        # SELECT otherwise takes and releases a file lock (~270 us measured),
        # and _read_source looks up one shared string per text cell outside
        # any write transaction here (build_index has its INSERTs open).
        self.db.execute("PRAGMA locking_mode=EXCLUSIVE")
        self.db.execute("PRAGMA synchronous=OFF")
        self.sink = sink

    def put(self, node, *, only_if_absent=False):
        self.sink.put(node)

    # The product's reader also feeds its own grids (index schema 5); the
    # prototype keeps its sink instead.
    def begin_grid(self, *, epoch1904):
        pass

    def note_array(self, sheet, r1, c1, r2, c2):
        pass

    def finish_grid(self):
        pass


def build_store(workbook: Path, store_path: Path) -> dict:
    """Stream the workbook with serve's reader into a compact value store."""
    from linexcel.lazy_store import _read_source

    started = time.perf_counter()
    store_path.unlink(missing_ok=True)
    scratch = store_path.with_suffix(".scratch.sqlite")
    scratch.unlink(missing_ok=True)
    db = sqlite3.connect(store_path)
    db.executescript("""
        PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF;
        CREATE TABLE sheets(name TEXT PRIMARY KEY, grid TEXT, width INTEGER,
            literal_cells INTEGER, formula_cells INTEGER);
        CREATE TABLE wide(sheet TEXT, row INTEGER, col INTEGER, value);
        CREATE TABLE exceptions(sheet TEXT, row INTEGER, col INTEGER, kind TEXT,
            value_json TEXT);
        CREATE TABLE columns(sheet TEXT, col INTEGER, stats_json TEXT,
            PRIMARY KEY(sheet, col));
        CREATE TABLE regions(id TEXT PRIMARY KEY, sheet TEXT, r1 INTEGER, c1 INTEGER,
            r2 INTEGER, c2 INTEGER, data_json TEXT);
        CREATE TABLE tables(name TEXT, sheet TEXT, data_json TEXT);
        CREATE TABLE source(key TEXT PRIMARY KEY, value TEXT);
    """)
    sink = _Sink(db)
    capture = _CaptureStore(scratch, sink)
    metadata, _names, _links = _read_source(workbook, capture)
    sink.close()
    db.executemany(
        "INSERT INTO tables VALUES (?,?,?)",
        capture.db.execute(
            "SELECT name_key, sheet, data FROM table_definitions"
        ).fetchall(),
    )
    capture.db.close()
    scratch.unlink(missing_ok=True)
    db.execute("CREATE INDEX exceptions_position ON exceptions(sheet, row, col)")
    from linexcel.lazy_store import file_sha256

    source = {
        "workbookSha256": file_sha256(workbook),
        "sheets": metadata["sheets"],
        "sheetStatus": {s["name"]: s["status"] for s in metadata["sheetDetails"]},
        "diagnostics": metadata["diagnostics"],
        "dateSystem": metadata["dateSystem"],
    }
    db.executemany(
        "INSERT INTO source VALUES (?,?)", [(k, _dump(v)) for k, v in source.items()]
    )
    db.commit()
    db.close()
    return {
        "seconds": round(time.perf_counter() - started, 2),
        "storeMB": round(store_path.stat().st_size / 1e6, 1),
    }


def _header_row(db, grid, r1, r2, c1, c2):
    """First row (within 20) whose non-empty cells are all text and cover half
    the columns, followed by a row holding at least one non-text value."""
    cols = ",".join(f"c{c}" for c in range(c1, c2 + 1))
    rows = db.execute(
        f'SELECT row,{cols} FROM "{grid}" '
        "WHERE row BETWEEN ? AND ? ORDER BY row LIMIT ?",
        (r1, r2, HEADER_SCAN_ROWS + 1),
    ).fetchall()
    for index, (row, *values) in enumerate(rows[:-1]):
        present = [v for v in values if v is not None]
        if len(present) * 2 >= len(values) and all(isinstance(v, str) for v in present):
            following = rows[index + 1][1:]
            if any(v is not None and not isinstance(v, str) for v in following):
                return row, [str(v) if v is not None else None for v in values]
    return None, None


def detect_regions(store_path: Path) -> list[dict]:
    """Excel tables first (formula columns split off), then runs of
    formula-free columns. Mixed columns stay per-cell nodes."""
    db = sqlite3.connect(store_path)
    db.execute("DELETE FROM regions")
    regions = []
    sheets = {
        name: (grid, width)
        for name, grid, width in db.execute("SELECT name, grid, width FROM sheets")
    }
    stats = {}
    for sheet, col, data in db.execute("SELECT sheet, col, stats_json FROM columns"):
        stats.setdefault(sheet, {})[col] = json.loads(data)
    claimed: dict[str, list] = {}

    def add(sheet, r1, c1, r2, c2, kind, header_row, headers, **extra):
        cells = sum(
            min(r2, s["lastRow"]) - max(r1, s["firstRow"]) + 1
            for c, s in stats[sheet].items()
            if c1 <= c <= c2 and s["firstRow"] is not None
        )
        region = {
            "id": f"data:{sheet}!R{r1}C{c1}:R{r2}C{c2}",
            "kind": kind,
            "sheet": sheet,
            "bounds": [r1, c1, r2, c2],
            "headerRow": header_row,
            "headers": headers,
            "rowCount": r2 - r1 + 1,
            "columnCount": c2 - c1 + 1,
            "approxCells": cells,
            **extra,
        }
        regions.append(region)
        claimed.setdefault(sheet, []).append((r1, c1, r2, c2))

    for _name, sheet, data in db.execute(
        "SELECT name, sheet, data_json FROM tables"
    ).fetchall():
        table = json.loads(data)
        if sheet not in sheets:
            continue
        r1, c1, r2, c2 = table["bounds"]
        body = (r1 + table.get("headerRows", 1), r2 - table.get("totalRows", 0))
        formula_cols = [
            c
            for c in range(c1, c2 + 1)
            if stats[sheet].get(c, {}).get("formulaCells", 0)
        ]
        run = []
        for c in [*range(c1, c2 + 1), None]:
            if c is not None and c not in formula_cols:
                run.append(c)
                continue
            if run and (body[1] - body[0] + 1) * len(run) >= REGION_MIN_CELLS:
                headers = [table["headers"][x - c1] for x in run]
                add(
                    sheet,
                    body[0],
                    run[0],
                    body[1],
                    run[-1],
                    "excel_table",
                    r1 if table.get("headerRows", 1) else None,
                    headers,
                    table=table["name"],
                    excludedFormulaColumns=[
                        table["headers"][x - c1] for x in formula_cols
                    ],
                )
            run = []
    for sheet, (grid, width) in sheets.items():
        columns = stats.get(sheet, {})
        literal = sorted(
            c
            for c, s in columns.items()
            if s["formulaCells"] == 0 and s["literalCells"]
        )
        runs, run = [], []
        for c in literal:
            if run and c != run[-1] + 1:
                runs.append(run)
                run = []
            run.append(c)
        if run:
            runs.append(run)
        for run in runs:
            r1 = min(columns[c]["firstRow"] for c in run)
            r2 = max(columns[c]["lastRow"] for c in run)
            c1, c2 = run[0], run[-1]
            if any(
                not (r2 < a or b < r1 or c2 < cc or d < c1)
                for a, cc, b, d in claimed.get(sheet, [])
            ):
                continue  # inside an Excel table already handled
            if (r2 - r1 + 1) * (c2 - c1 + 1) < REGION_MIN_CELLS or c2 > GRID_COLUMNS:
                continue
            header_row, headers = _header_row(db, grid, r1, r2, c1, c2)
            data_r1 = header_row + 1 if header_row else r1
            add(
                sheet,
                data_r1,
                c1,
                r2,
                c2,
                "formula_free_columns",
                header_row,
                headers,
                preambleRows=[r1, header_row - 1]
                if header_row and header_row > r1
                else None,
            )
    for region in regions:
        _describe(db, sheets[region["sheet"]][0], region, stats[region["sheet"]])
        db.execute(
            "INSERT INTO regions VALUES (?,?,?,?,?,?,?)",
            (region["id"], region["sheet"], *region["bounds"], _dump(region)),
        )
    db.commit()
    db.close()
    return regions


def _describe(db, grid, region, stats):
    """Schema, formats, exceptions, fingerprint and a bounded preview."""
    r1, c1, r2, c2 = region["bounds"]
    sheet = region["sheet"]
    started = time.perf_counter()
    cols = ",".join(f"c{c}" for c in range(c1, c2 + 1))
    digest = hashlib.sha256()
    types = [Counter() for _ in range(c1, c2 + 1)]
    count = 0
    for row in db.execute(
        f'SELECT row,{cols} FROM "{grid}" WHERE row BETWEEN ? AND ? ORDER BY row',
        (r1, r2),
    ):
        digest.update(_dump(row).encode())
        for index, value in enumerate(row[1:]):
            if value is None:
                types[index]["empty"] += 1
            else:
                count += 1
                types[index][type(value).__name__] += 1
    exceptions = Counter(
        dict(
            db.execute(
                "SELECT kind, count(*) FROM exceptions "
                "WHERE sheet=? AND row BETWEEN ? AND ? "
                "AND col BETWEEN ? AND ? GROUP BY kind",
                (sheet, r1, r2, c1, c2),
            ).fetchall()
        )
    )
    for kind, value in db.execute(
        "SELECT kind, value_json FROM exceptions WHERE sheet=? AND row BETWEEN ? AND ? "
        "AND col BETWEEN ? AND ? "
        "AND kind IN ('bool','error','iso_date','json','unreadable') ORDER BY row, col",
        (sheet, r1, r2, c1, c2),
    ):
        digest.update(f"{kind}:{value}".encode())
    schema = []
    for offset, c in enumerate(range(c1, c2 + 1)):
        observed = {k: v for k, v in types[offset].items() if k != "empty"}
        dtype = (
            "empty"
            if not observed
            else "number"
            if set(observed) <= {"int", "float"}
            else "text"
            if set(observed) == {"str"}
            else "mixed"
        )
        formats = (stats.get(c) or {}).get("formats") or {}
        if dtype == "number" and any(_is_date(f) for f in formats):
            dtype = (
                "date_serial"
                if all(_is_date(f) for f in formats)
                else "mixed_number_date"
            )
        schema.append(
            {
                "column": c,
                "name": (region["headers"] or [None] * (c2 - c1 + 1))[offset],
                "type": dtype,
                "observed": observed,
                "empty": types[offset]["empty"],
                "formats": formats,
                "min": (stats.get(c) or {}).get("min"),
                "max": (stats.get(c) or {}).get("max"),
            }
        )
    preview_columns = ",".join(f"c{c}" for c in range(c1, min(c2, c1 + 11) + 1))
    preview = db.execute(
        f'SELECT row,{preview_columns} FROM "{grid}" '
        "WHERE row BETWEEN ? AND ? ORDER BY row LIMIT 10",
        (r1, r2),
    ).fetchall()
    region.update(
        cells=count,
        schema=schema,
        exceptions=dict(exceptions),
        calculable=not exceptions.get("unreadable"),
        fingerprint="sha256:" + digest.hexdigest(),
        preview={"rows": [list(r) for r in preview], "bounded": "10 rows x 12 columns"},
        describeSeconds=round(time.perf_counter() - started, 2),
    )


def _is_date(number_format: str) -> bool:
    from openpyxl.styles.numbers import is_date_format

    return is_date_format(number_format)


def regions_for(store_path: Path, sheet: str) -> list[tuple]:
    db = sqlite3.connect(store_path)
    rows = db.execute(
        "SELECT id, r1, c1, r2, c2 FROM regions WHERE sheet=?", (sheet,)
    ).fetchall()
    db.close()
    return rows


def link(store_path: Path, index_path: Path, sample: int = 3) -> list[dict]:
    """Formulas reading a region point to the region with the sub-range read."""
    db = sqlite3.connect(store_path)
    regions = [
        (rid, sheet, (r1, c1, r2, c2))
        for rid, sheet, r1, c1, r2, c2 in db.execute(
            "SELECT id, sheet, r1, c1, r2, c2 FROM regions"
        )
    ]
    db.close()
    index = sqlite3.connect(f"file:{index_path}?mode=ro", uri=True)
    from linexcel.lazy_store import _unpack

    links = []
    for seq, (rid, data) in enumerate(
        index.execute("SELECT id, data FROM nodes WHERE kind='range'")
    ):
        node = _unpack(data)
        bounds = node.get("bounds")
        if not bounds:
            continue
        for region_id, sheet, (r1, c1, r2, c2) in regions:
            if node.get("sheet") != sheet:
                continue
            a, b, c, d = bounds
            sub = (max(a, r1), max(b, c1), min(c, r2), min(d, c2))
            if sub[0] > sub[2] or sub[1] > sub[3]:
                continue
            consumers = [
                t
                for (t,) in index.execute(
                    "SELECT target FROM edges WHERE source=? LIMIT 1000", (rid,)
                )
            ]
            links.append(
                {
                    "range": rid,
                    "region": region_id,
                    "subRange": list(sub),
                    "coversWholeRange": tuple(bounds) == sub,
                    "consumers": len(consumers),
                    "consumerSample": consumers[:sample],
                }
            )
    direct = Counter()
    for region_id, sheet, (r1, c1, r2, c2) in regions:
        direct[region_id] = index.execute(
            "SELECT count(*) FROM edges e JOIN nodes n ON n.id=e.source "
            "WHERE n.sheet=? AND n.kind='input' "
            "AND n.row BETWEEN ? AND ? AND n.col BETWEEN ? AND ?",
            (sheet, r1, r2, c1, c2),
        ).fetchone()[0]
    index.close()
    return [{"directCellReferences": dict(direct)}, *links]


def window(
    store_path: Path,
    sheet: str,
    row: int,
    column: int,
    rows: int = 60,
    columns: int = 20,
) -> dict:
    """One presentation window served from the store: O(window), not O(sheet)."""
    started = time.perf_counter()
    db = sqlite3.connect(f"file:{store_path}?mode=ro", uri=True)
    grid, width = db.execute(
        "SELECT grid, width FROM sheets WHERE name=?", (sheet,)
    ).fetchone()
    last = min(width, column + columns - 1)
    cols = ",".join(f"c{c}" for c in range(column, last + 1)) or "NULL"
    data = db.execute(
        f'SELECT row,{cols} FROM "{grid}" WHERE row BETWEEN ? AND ? ORDER BY row',
        (row, row + rows - 1),
    ).fetchall()
    exceptions = db.execute(
        "SELECT row, col, kind, value_json FROM exceptions "
        "WHERE sheet=? AND row BETWEEN ? AND ? AND col BETWEEN ? AND ?",
        (sheet, row, row + rows - 1, column, last),
    ).fetchall()
    db.close()
    return {
        "rows": len(data),
        "cells": sum(v is not None for r in data for v in r[1:]),
        "exceptions": len(exceptions),
        "milliseconds": round((time.perf_counter() - started) * 1000, 2),
    }


def region_values(
    store_path: Path, sheet: str, bounds
) -> tuple[list[list], list[tuple]]:
    """Engine values for a sub-range, row-major, with exception overrides.
    Returns (rows, unreadable cells). Raises nothing silently."""
    r1, c1, r2, c2 = bounds
    db = sqlite3.connect(f"file:{store_path}?mode=ro", uri=True)
    grid = db.execute("SELECT grid FROM sheets WHERE name=?", (sheet,)).fetchone()[0]
    cols = ",".join(f"c{c}" for c in range(c1, c2 + 1))
    data = {
        row: list(values)
        for row, *values in db.execute(
            f'SELECT row,{cols} FROM "{grid}" WHERE row BETWEEN ? AND ?', (r1, r2)
        )
    }
    unreadable = []
    for row, col, kind, value in db.execute(
        "SELECT row, col, kind, value_json FROM exceptions "
        "WHERE sheet=? AND row BETWEEN ? AND ? "
        "AND col BETWEEN ? AND ? AND kind IN ('bool','error','json','unreadable')",
        (sheet, r1, r2, c1, c2),
    ):
        if kind == "unreadable":
            unreadable.append((row, col))
            continue
        decoded = json.loads(value)
        if kind == "error":
            from linexcel.lazy import _INPUT_ERROR_KINDS

            decoded = {"type": "Error", "kind": _INPUT_ERROR_KINDS[decoded]}
        data.setdefault(row, [None] * (c2 - c1 + 1))[col - c1] = decoded
    db.close()
    rows = [data.get(r, [None] * (c2 - c1 + 1)) for r in range(r1, r2 + 1)]
    return rows, unreadable
