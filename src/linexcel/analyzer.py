"""Builds the lineage graph of an Excel workbook.

Steps:
1. structure (sheets, dimensions, defined names) via openpyxl read-only;
2. formulas + computed values via the Rust engine formualizer;
3. grouping of stretched formulas by R1C1 canonicalization —
   a column of 50,000 copied formulas becomes ONE node;
4. resolution of precedents (cells, ranges, names, other sheets);
5. decomposition of each composite formula into individually evaluated steps
   in a scratch sheet of the engine;
6. lineage of extracted VBA code (oletools).
"""

from __future__ import annotations

import datetime
import io
import itertools
import re
import sys
import time
import uuid
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

import formualizer as fz
from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format

from linexcel.refs import (
    Rect,
    num_to_col,
    parse_ref_detailed,
    stretch_ref,
)
from linexcel.rewrite import canonical_r1c1, qualify_sheet
from linexcel.vba import VbaProc, analyze_vba, extract_vba_modules

# Guards to stay responsive on large workbooks.
SCAN_CHUNK_ROWS = 20_000
MAX_CELLS_PER_SHEET = 4_000_000
SMALL_RANGE_CELLS = 20_000
MAX_NODES_PER_SHEET = 400
MAX_STEPS_PER_FORMULA = 48
MAX_SCRATCH_EVALS = 4_000
MAX_VALUE_SAMPLE = 5
MAX_VBA_CODE_CHARS = 6_000
MAX_VALUE_WARNINGS = 25
# Chained recovery: how deep the precedent walk goes, and how wide a referenced
# range may be before its cells are left to the engine. Both only bound the
# work; a cell that is skipped simply keeps the value the engine reports.
MAX_CHAIN_DEPTH = 24
MAX_CHAIN_RANGE_CELLS = 4_096
SCRATCH_SHEET = "__lineage_scratch__"
# Written into the scratch cell before each guarded evaluation: when the engine
# fails to compute an expression it silently keeps the previous cell value
# instead of raising, so an unchanged marker is how we detect that failure.
SCRATCH_SENTINEL = "__linexcel_no_value__"
EXCEL_EPOCH_1900 = datetime.datetime(1899, 12, 30)
EXCEL_EPOCH_1904 = datetime.datetime(1904, 1, 1)
# Serials 1..59 (Jan/Feb 1900) sit before Excel's phantom 1900-02-29, so their
# real epoch is 1899-12-31: serial 1 → 1900-01-01. Serial 60 is the phantom
# day itself and matches no real date. From 61 on, the 1899-12-30 epoch
# already absorbs the phantom day.
EPOCH_EARLY_1900 = datetime.datetime(1899, 12, 31)
GUARD_FUNCTIONS = {"IFERROR", "IFNA"}
#: Engine error kinds → the text a spreadsheet shows for them. These are values
#: a cell genuinely holds: openpyxl reads the stored one back as this same text,
#: so the graph carries the text on both sides and they compare directly.
ERROR_KIND_TEXT = {
    "Div": "#DIV/0!",
    "Na": "#N/A",
    "Name": "#NAME?",
    "Num": "#NUM!",
    "Null": "#NULL!",
    "Ref": "#REF!",
    "Value": "#VALUE!",
    "Spill": "#SPILL!",
    "Calc": "#CALC!",
}
EXCEL_ERRORS = frozenset(ERROR_KIND_TEXT.values())
#: Error kinds that are *not* a cell value: the engine reporting that it could
#: not compute, rather than a result. ``NImpl`` is a function or operator it does
#: not implement — the range intersection in ``=SUM(D2:D10 D5:D20)``, say — and
#: ``Circ`` a reference cycle, which Excel resolves by convention rather than by
#: computing. Neither may be shown as "the value linexcel recalculated", because
#: linexcel recalculated nothing; the cell falls back to what the file stores.
UNCOMPUTED_ERROR_KINDS = frozenset({"NImpl", "Cancelled", "Circ"})
#: How many uncomputable cells the warning names before it stops listing them.
MAX_UNCOMPUTED_LISTED = 10


@dataclass
class FormulaGroup:
    """A set of cells on a sheet sharing the same R1C1 formula."""

    sheet: str
    r1c1: str
    cells: list[tuple[int, int]] = field(default_factory=list)
    formulas: dict[tuple[int, int], str] = field(default_factory=dict)

    @property
    def rep(self) -> tuple[int, int]:
        return min(self.cells)

    @property
    def bbox(self) -> tuple[int, int, int, int]:
        rows = [r for r, _ in self.cells]
        cols = [c for _, c in self.cells]
        return min(rows), min(cols), max(rows), max(cols)


class _Budget:
    def __init__(self, limit: int):
        self.left = limit

    def take(self) -> bool:
        if self.left <= 0:
            return False
        self.left -= 1
        return True


class CachedValues:
    """Values the spreadsheet application cached in the file.

    They are what the user last saw on screen. Workbooks written by openpyxl
    carry no cache for formula cells, workbooks saved by Excel or LibreOffice
    do; either way constants and their number formats are always readable.
    """

    def __init__(
        self,
        values: dict[tuple[str, int, int], Any],
        date_cells: set[tuple[str, int, int]],
        epoch_1904: bool,
    ):
        self._values = values
        self._date_cells = date_cells
        self.epoch_1904 = epoch_1904

    def get(self, sheet: str, row: int, col: int) -> Any:
        return self._values.get((sheet, row, col))

    def is_date(self, sheet: str, row: int, col: int) -> bool:
        return (sheet, row, col) in self._date_cells

    def __len__(self) -> int:
        return len(self._values)


def load_cached_values(data: bytes) -> CachedValues:
    """Read the file's cached values once, keyed by (sheet, row, col).

    python-calamine (the Rust engine) is the hot path: it returns native Python
    types directly, so dates are detected by type instead of by number_format
    string, and it is roughly an order of magnitude faster than openpyxl on
    large files. openpyxl remains the fallback for the rare file calamine cannot
    open; there its number_format-based date detection keeps the edge case (a
    number formatted as a date but stored as a float) covered.
    """
    try:
        return _load_cached_values_calamine(data)
    except Exception:
        return _load_cached_values_openpyxl(data)


def _load_cached_values_calamine(data: bytes) -> CachedValues:
    """Fast path: read cached values via python-calamine.

    ``to_python(skip_empty_area=False)`` preserves the cell coordinates —
    linexcel keys values by ``(sheet, row, col)`` 1-indexed, and the default
    ``skip_empty_area=True`` would shift everything toward the top-left.
    Empty cells come back as ``""`` (calamine) or ``None``; both are skipped,
    matching openpyxl which reads an empty cell back as ``None``.

    Dates are detected by type: calamine returns ``datetime.date`` for
    date-only cells and ``datetime.datetime`` for timestamps. openpyxl always
    returns ``datetime.datetime`` (even for a date-only cell), so a bare
    ``datetime.date`` is normalized to a midnight ``datetime.datetime`` to keep
    downstream equality checks and the resolver seeing one consistent type.
    """
    from python_calamine import CalamineWorkbook

    values: dict[tuple[str, int, int], Any] = {}
    date_cells: set[tuple[str, int, int]] = set()
    epoch_1904 = _detect_epoch_1904(data)
    wb = CalamineWorkbook.from_object(io.BytesIO(data))
    for name in wb.sheet_names:
        sheet = wb.get_sheet_by_name(name)
        rows = sheet.to_python(skip_empty_area=False)
        scanned = 0
        for r_idx, row in enumerate(rows):
            scanned += len(row)
            if scanned > MAX_CELLS_PER_SHEET:
                break
            for c_idx, v in enumerate(row):
                if v is None or v == "":
                    continue
                key = (name, r_idx + 1, c_idx + 1)
                if isinstance(v, datetime.datetime):
                    date_cells.add(key)
                elif isinstance(v, datetime.date):
                    # openpyxl reads a date-only cell as a midnight datetime;
                    # normalize so the two readers are interchangeable.
                    v = datetime.datetime(v.year, v.month, v.day)
                    date_cells.add(key)
                values[key] = v
    return CachedValues(values, date_cells, epoch_1904)


def _load_cached_values_openpyxl(data: bytes) -> CachedValues:
    """Fallback: read cached values via openpyxl with number_format date detection."""
    values: dict[tuple[str, int, int], Any] = {}
    date_cells: set[tuple[str, int, int]] = set()
    epoch_1904 = False
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        return CachedValues(values, date_cells, epoch_1904)
    try:
        epoch_1904 = getattr(wb.epoch, "year", 1899) == 1904
        for ws in wb.worksheets:
            scanned = 0
            for row in ws.iter_rows():
                scanned += len(row)
                if scanned > MAX_CELLS_PER_SHEET:
                    break
                for cell in row:
                    # read-only sheets pad gaps with EmptyCell (no coordinates)
                    r = getattr(cell, "row", None)
                    c = getattr(cell, "column", None)
                    if r is None or c is None:
                        continue
                    key = (ws.title, r, c)
                    if _is_date_format(getattr(cell, "number_format", None)):
                        date_cells.add(key)
                    if cell.value is not None:
                        values[key] = cell.value
    except Exception:
        pass
    finally:
        wb.close()
    return CachedValues(values, date_cells, epoch_1904)


def _detect_epoch_1904(data: bytes) -> bool:
    """True if the workbook declares the 1904 date system.

    calamine converts 1904-epoch dates to correct values itself, but linexcel
    also needs the flag to interpret *engine* serials (from formualizer, not
    calamine) via :func:`serial_to_date_text`. The flag lives on
    ``<workbookPr date1904="1" />`` in ``xl/workbook.xml``; reading it from the
    zip avoids opening openpyxl a second time just for this one attribute.
    """
    import re
    import zipfile

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    except Exception:
        return False
    return bool(re.search(r"""date1904=["']1""", xml))


class _ValueResolver:
    """Single entry point for reading the value of a cell.

    The engine is all-or-nothing: one cell pointing at a missing sheet makes
    ``evaluate_all`` raise, and from then on every formula cell reads back as
    ``None`` — the whole workbook loses its values because of one bad
    reference. So each read is resolved in order: engine value, per-cell
    recalculation, isolated re-evaluation of the formula in the scratch sheet
    (with the IFERROR/IFNA fallback branch when the guarded expression itself
    cannot be computed), and finally the value cached in the file.

    A scratch evaluation only sees constants: a reference to another *formula*
    cell reads back as blank, so ``=A3+A4`` silently yielded 0 and
    ``=SUM(A1:A7)`` counted the constants only. Recovery therefore resolves the
    precedents of a formula first, deepest one first, and feeds each recovered
    value back into the engine as a constant so the next evaluation — direct
    reference or range — reads a number instead of a formula it cannot compute.
    """

    def __init__(
        self,
        engine,
        engine_sheets: set[str],
        cached: CachedValues,
        warnings: list[str],
        budget: _Budget,
        scratch_ready: bool,
        engine_alive: bool = True,
        sheet_dims: dict[str, tuple[int, int]] | None = None,
    ):
        self.engine = engine
        self.engine_sheets = engine_sheets
        self.cached = cached
        self.warnings = warnings
        self.budget = budget
        self.scratch_ready = scratch_ready
        self.sheet_dims = sheet_dims or {}
        self._engine_alive = engine_alive
        self._compared: set[tuple[str, int, int]] = set()
        self._n_mismatches = 0
        self._resolved: dict[tuple[str, int, int], tuple[Any, str | None]] = {}
        self._resolving: set[tuple[str, int, int]] = set()
        self._uncomputed: list[str] = []
        self.n_recovered = 0
        self.n_unrecovered = 0
        self._step_cache: dict[str, tuple[Any, bool]] = {}

    # -- public API --------------------------------------------------------
    def value(
        self, sheet: str | None, row: int, col: int, formula: str | None = None
    ) -> tuple[Any, str | None, str | None]:
        """Return ``(value, source, date_text)`` for one cell."""
        if sheet is None:
            return None, None, None
        if sheet not in self.engine_sheets:
            return self._from_cache(sheet, row, col)
        raw, source = self._engine_read(sheet, row, col, formula)
        if _is_uncomputed(raw):
            self._note_uncomputed(sheet, row, col)
            raw = None
        if raw is None:
            return self._from_cache(sheet, row, col)
        date_text = self._date_text(sheet, row, col, raw)
        self._check_mismatch(sheet, row, col, raw, date_text)
        return _jsonable(raw), source, date_text

    def describe(
        self, sheet: str | None, row: int, col: int, formula: str | None = None
    ) -> dict[str, Any]:
        """Graph fields for a cell: value, provenance, cache and date."""
        value, source, date_text = self.value(sheet, row, col, formula)
        fields: dict[str, Any] = {"value": value}
        if source is not None:
            fields["valueSource"] = source
        cached = self.cached_value(sheet, row, col)
        if cached is not None:
            fields["cachedValue"] = cached
        if date_text is not None:
            fields["valueDate"] = date_text
        return fields

    def cached_value(self, sheet: str | None, row: int, col: int) -> Any:
        if sheet is None:
            return None
        raw = self.cached.get(sheet, row, col)
        # dates stay dates in the graph: a midnight datetime serializes as the
        # bare ``YYYY-MM-DD`` so it compares cleanly against ``valueDate``
        if isinstance(raw, datetime.datetime):
            if raw.time() == datetime.time.min:
                return raw.date().isoformat()
            return raw.isoformat(sep=" ")
        if isinstance(raw, datetime.date):
            return raw.isoformat()
        return _jsonable(raw)

    def eval_expr(self, expr: str, sheet: str) -> tuple[Any, bool]:
        """Evaluate an expression in the scratch sheet, if budget allows.

        A step the engine cannot compute reads back as *not evaluated* rather
        than as an error: "#NULL!" under a step is the spreadsheet's own verdict
        on the formula, and the engine hitting its own limit is not that.
        """
        raw, ok = self._eval_raw(expr, sheet)
        if _is_uncomputed(raw):
            return None, False
        return _jsonable(raw), ok

    # -- internals ---------------------------------------------------------
    def _eval_raw(self, expr: str, sheet: str) -> tuple[Any, bool]:
        """Scratch evaluation keeping the engine value as-is.

        Errors come back as ``{"type": "Error", ...}``; they only become the
        string the graph carries at the very end, because an error fed back
        into the engine is what lets a guard downstream still fire.
        """
        if not self.scratch_ready or not self.budget.take():
            return None, False
        cached = self._step_cache.pop(expr, None)
        if cached is not None:
            return cached
        return _scratch_eval(self.engine, expr, sheet)

    def preload_steps(self, exprs: list[str], sheet: str) -> None:
        """Batch-evaluate step expressions in one ``evaluate_cells`` call.

        Each ``evaluate_cell`` on a large workbook recalculates the whole
        dependency graph (~77 ms).  Batching N expressions into a single
        ``evaluate_cells`` pays that cost once instead of N times.  When
        ``engine_alive`` is False the engine state mutates during
        decomposition (``_remember`` calls ``set_value``), so batching is
        unsafe — the pre-computed values would be stale.  In that case the
        cache stays empty and every expression falls back to
        ``_scratch_eval`` one at a time, exactly as before.
        """
        self._step_cache.clear()
        if not self.scratch_ready or not self._engine_alive or not exprs:
            return
        # ponytail: dedup preserves order — identical sub-expressions share
        # one scratch cell and one cache entry; _eval_raw pops on first hit
        # so the second occurrence falls through to _scratch_eval, matching
        # the old per-call behaviour.
        seen: set[str] = set()
        unique: list[str] = []
        for e in exprs:
            if e not in seen:
                seen.add(e)
                unique.append(e)
        targets: list[tuple[str, int, int]] = []
        valid: list[str] = []
        for i, e in enumerate(unique):
            try:
                qualified = qualify_sheet(e, sheet)
            except Exception:
                continue
            try:
                self.engine.set_formula(SCRATCH_SHEET, 2, i + 1, qualified)
            except Exception:
                continue
            targets.append((SCRATCH_SHEET, 2, i + 1))
            valid.append(e)
        if not targets:
            return
        try:
            results = self.engine.evaluate_cells(targets)
        except Exception:
            return  # a broken reference poisons the batch — fall back
        for e, val in zip(valid, results):
            if val is not None and not _is_uncomputed(val):
                self._step_cache[e] = (_jsonable(val), True)
            else:
                self._step_cache[e] = (None, False)

    def _engine_read(
        self, sheet: str, row: int, col: int, formula: str | None
    ) -> tuple[Any, str | None]:
        memo = self._resolved.get((sheet, row, col))
        if memo is not None:
            return memo
        try:
            raw = self.engine.get_value(sheet, row, col)
        except Exception:
            raw = None
        if raw is not None:
            return raw, "engine"
        if formula is None:
            formula = self._formula_at(sheet, row, col)
        if not formula:
            return None, None
        return self._recover(sheet, row, col, formula)

    def _recover(
        self, sheet: str, row: int, col: int, formula: str
    ) -> tuple[Any, str | None]:
        uncomputed = None
        if self._engine_alive:
            try:
                raw = self.engine.evaluate_cell(sheet, row, col)
                if raw is not None:
                    if _is_uncomputed(raw):
                        uncomputed = raw
                    else:
                        return raw, "engine"
            except Exception:
                # The first failure poisons the engine for good: every later
                # whole-graph evaluation would raise the same way.
                self._engine_alive = False
        expr = formula if formula.startswith("=") else "=" + formula
        result = self._eval_formula(sheet, expr, 0)
        if result[0] is None and uncomputed is not None:
            result = uncomputed, None
        return self._remember(sheet, row, col, result)

    def _eval_formula(
        self, sheet: str, expr: str, depth: int
    ) -> tuple[Any, str | None]:
        """Evaluate one formula on its own, precedents resolved first.

        An expression the engine cannot compute counts as a failure, not as a
        result: the IFERROR/IFNA fallback branch is then tried, exactly as it is
        when the evaluation itself does not come back.
        """
        if not self._engine_alive:
            self._resolve_precedents(sheet, expr, depth)
        raw, ok = self._eval_raw(expr, sheet)
        uncomputed = raw if ok and raw is not None and _is_uncomputed(raw) else None
        if ok and raw is not None and not _is_uncomputed(raw):
            return raw, "engine"
        fallback = _guard_fallback_expr(expr)
        if fallback is not None:
            raw, ok = self._eval_raw(fallback, sheet)
            if uncomputed is None and ok and raw is not None and _is_uncomputed(raw):
                uncomputed = raw
            if ok and raw is not None and not _is_uncomputed(raw):
                return raw, "fallback"
        if uncomputed is not None:
            return uncomputed, None
        return None, None

    def _resolve_precedents(self, sheet: str, expr: str, depth: int) -> None:
        """Recover every formula cell the expression reads, deepest first."""
        if depth >= MAX_CHAIN_DEPTH:
            return
        try:
            ast_dict = fz.parse(expr).to_dict()
        except Exception:
            return
        for ref in _collect_ref_strings(ast_dict):
            detail = parse_ref_detailed(ref, default_sheet=sheet)
            if detail is None or detail.rect.sheet not in self.engine_sheets:
                continue
            rect = detail.rect
            dims = self.sheet_dims.get(rect.sheet or "")
            if dims is not None:
                clipped = rect.clipped(*dims)
                if clipped is None:
                    continue
                rect = clipped
            if rect.ncells > MAX_CHAIN_RANGE_CELLS or rect.sheet is None:
                continue
            for r in range(rect.r1, rect.r2 + 1):
                for c in range(rect.c1, rect.c2 + 1):
                    self._resolve_chain(rect.sheet, r, c, depth + 1)

    def _resolve_chain(self, sheet: str, row: int, col: int, depth: int) -> None:
        key = (sheet, row, col)
        # ``_resolving`` breaks the self-referencing formula (=B1+B2 in B2):
        # the cell stays unresolved and reads as blank, as it did before.
        if key in self._resolved or key in self._resolving:
            return
        if depth >= MAX_CHAIN_DEPTH or self.budget.left <= 0:
            return
        try:
            if self.engine.get_value(sheet, row, col) is not None:
                return  # constant: the engine already resolves it on its own
        except Exception:
            pass
        formula = self._formula_at(sheet, row, col)
        if not formula:
            return
        expr = formula if formula.startswith("=") else "=" + formula
        self._resolving.add(key)
        try:
            result = self._eval_formula(sheet, expr, depth)
        finally:
            self._resolving.discard(key)
        self._remember(sheet, row, col, result)

    def _remember(
        self, sheet: str, row: int, col: int, result: tuple[Any, str | None]
    ) -> tuple[Any, str | None]:
        """Memoize a recovered value and feed it back into the engine."""
        self._resolved[(sheet, row, col)] = result
        raw, _source = result
        if raw is None or _is_uncomputed(raw):
            self.n_unrecovered += 1
            return result
        self.n_recovered += 1
        if not self._engine_alive:
            # Only ever done on the engine rebuilt after a failed global
            # evaluation, which holds no computed value anyway. It drops the
            # formula of the cell, hence the memo read first in ``_engine_read``.
            try:
                self.engine.set_value(sheet, row, col, raw)
            except Exception:
                pass
        return result

    def _formula_at(self, sheet: str, row: int, col: int) -> str | None:
        try:
            return self.engine.get_formula(sheet, row, col)
        except Exception:
            return None

    def _from_cache(
        self, sheet: str, row: int, col: int
    ) -> tuple[Any, str | None, str | None]:
        raw = self.cached.get(sheet, row, col)
        if raw is None:
            return None, None, None
        date_text = _date_text_of(raw)
        return _jsonable(raw), "file", date_text

    def _date_text(self, sheet: str, row: int, col: int, raw: Any) -> str | None:
        if isinstance(raw, bool) or not isinstance(raw, (int, float)):
            return _date_text_of(raw)
        cached = self.cached.get(sheet, row, col)
        is_date = isinstance(cached, (datetime.datetime, datetime.date))
        if not is_date and not self.cached.is_date(sheet, row, col):
            return None
        return serial_to_date_text(raw, self.cached.epoch_1904)

    def _check_mismatch(
        self, sheet: str, row: int, col: int, raw: Any, date_text: str | None
    ) -> None:
        key = (sheet, row, col)
        if key in self._compared:
            return
        self._compared.add(key)
        cached = self.cached.get(sheet, row, col)
        if cached is None or not _values_differ(raw, cached, date_text):
            return
        self._n_mismatches += 1
        if self._n_mismatches > MAX_VALUE_WARNINGS:
            if self._n_mismatches == MAX_VALUE_WARNINGS + 1:
                self.warnings.append(
                    "More recalculated values differ from the file "
                    f"(only the first {MAX_VALUE_WARNINGS} are listed)"
                )
            return
        self.warnings.append(
            f"{sheet}!{a1(row, col)}: recalculated {_fmt_value(raw)} "
            f"differs from file value {_fmt_value(cached)}"
        )

    def _note_uncomputed(self, sheet: str, row: int, col: int) -> None:
        """Record a cell the engine could not compute, once."""
        label = f"{sheet}!{a1(row, col)}"
        if label not in self._uncomputed:
            self._uncomputed.append(label)

    def uncomputed_warning(self) -> str | None:
        """One line naming the cells left to the value stored in the file.

        Silence would be the wrong report here: those cells show a figure the
        panel attributes to the file, and nothing else would say that linexcel
        met a formula it cannot evaluate.
        """
        if not self._uncomputed:
            return None
        listed = ", ".join(self._uncomputed[:MAX_UNCOMPUTED_LISTED])
        rest = len(self._uncomputed) - MAX_UNCOMPUTED_LISTED
        if rest > 0:
            listed += f", and {rest} more"
        return (
            f"{len(self._uncomputed)} cell(s) could not be computed by the engine "
            f"and keep the value stored in the file: {listed}"
        )


def a1(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def analyze_workbook(
    data: bytes,
    filename: str = "workbook.xlsx",
    *,
    verbose: bool = False,
) -> dict[str, Any]:
    """Full analysis: returns the JSON-serializable graph and the engine."""
    warnings: list[str] = []
    _t0 = time.perf_counter()

    def _v(label: str, t: float) -> None:
        if verbose:
            elapsed = time.perf_counter() - t
            print(f"[linexcel] {label}: {elapsed:.1f}s", file=sys.stderr)

    # --- 1. structure -----------------------------------------------------
    _t = time.perf_counter()
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

    # Table detection: openpyxl read-only mode does not expose tables, so the
    # workbook is opened once more in normal mode to read TableObjects. The
    # result is a per-cell lookup used to enrich lineage nodes.
    table_index = _build_table_index(data)
    _v("structure+tables", _t)

    # values the file itself carries: last resort, and the only source of
    # dates and of what the user actually saw on screen
    _t = time.perf_counter()
    cached = load_cached_values(data)
    _v("cached_values", _t)

    # --- 2. computation engine -------------------------------------------
    _t = time.perf_counter()
    engine = fz.Workbook.from_bytes(data)
    engine_sheets = set(engine.sheet_names)
    engine_alive = True
    quarantined: dict[tuple[str, int, int], str] = {}
    try:
        engine.evaluate_all()
    except Exception as exc:  # graph remains useful without values
        # A failed global evaluation does not just drop the values: the engine
        # then reports no formula at all, which would leave the graph empty.
        # Rebuilding from the bytes gives the formulas back.
        engine = fz.Workbook.from_bytes(data)
        # evaluate_all is all-or-nothing, and it gives up on the *first*
        # reference it cannot resolve — so a single formula pointing at another
        # workbook costs every other cell in the file its computed value. Set
        # those few cells aside and the pass usually completes, leaving only
        # them to the slower per-cell recovery.
        quarantined = _quarantine_unresolvable(engine, sheet_dims, engine_sheets)
        retried = False
        if quarantined:
            try:
                engine.evaluate_all()
                retried = True
            except Exception:
                engine = fz.Workbook.from_bytes(data)
        if retried:
            warnings.append(
                f"Global evaluation completed after isolating {len(quarantined)} "
                f"cell(s) whose references the engine cannot resolve; every other "
                f"cell was recomputed. Only those keep the value stored in the "
                f"file, if any. First blocker: {exc}"
            )
        else:
            warnings.append(f"Global evaluation incomplete: {exc}")
            # Values are recovered cell by cell further down.
            engine_alive = False
            quarantined = {}

    scratch_ready = _ensure_scratch(engine)
    _v("engine_init+evaluate_all", _t)
    budget = _Budget(MAX_SCRATCH_EVALS)
    resolver = _ValueResolver(
        engine,
        engine_sheets,
        cached,
        warnings,
        budget,
        scratch_ready,
        engine_alive=engine_alive,
        sheet_dims=sheet_dims,
    )

    # --- 3. extraction + grouping ----------------------------------------
    _t = time.perf_counter()
    groups: dict[tuple[str, str], FormulaGroup] = {}
    cell_owner: dict[str, dict[tuple[int, int], str]] = defaultdict(dict)
    formula_count = 0
    sheet_stats: list[dict[str, Any]] = []

    for sheet, (max_row, max_col) in sheet_dims.items():
        if sheet not in engine_sheets:
            warnings.append(f"Sheet '{sheet}' skipped (not loaded by engine)")
            continue
        n_formulas = 0
        scanned = 0
        fsheet = engine.sheet(sheet)
        for r0 in range(1, max_row + 1, SCAN_CHUNK_ROWS):
            r1 = min(r0 + SCAN_CHUNK_ROWS - 1, max_row)
            chunk_cells = (r1 - r0 + 1) * max_col
            if scanned + chunk_cells > MAX_CELLS_PER_SHEET:
                warnings.append(f"Sheet '{sheet}' truncated after {scanned:,} cells")
                break
            ra = fz.RangeAddress(sheet, r0, 1, r1, max_col)
            try:
                rows = fsheet.get_formulas(ra)
            except Exception as exc:
                warnings.append(f"Could not read formulas on {sheet}: {exc}")
                break
            scanned += (r1 - r0 + 1) * max_col
            for i, row_vals in enumerate(rows):
                r = r0 + i
                for j, f in enumerate(row_vals):
                    if not f:
                        # A quarantined cell reads back blank: its formula was
                        # removed so the rest of the workbook could evaluate.
                        f = quarantined.get((sheet, r, j + 1))
                    if not f:
                        continue
                    c = j + 1
                    n_formulas += 1
                    key = (sheet, canonical_r1c1(f, r, c))
                    grp = groups.get(key)
                    if grp is None:
                        grp = groups[key] = FormulaGroup(sheet, key[1])
                    grp.cells.append((r, c))
                    # row/col order scan: first cell seen is the representative
                    # (min), keep 3 example formulas
                    if len(grp.formulas) < 3:
                        grp.formulas[(r, c)] = f
        formula_count += n_formulas
        sheet_stats.append(
            {
                "name": sheet,
                "rows": max_row,
                "cols": max_col,
                "formulaCells": n_formulas,
            }
        )

    # --- 4. formula nodes -------------------------------------------------
    _v("extraction+grouping", _t)
    _t = time.perf_counter()
    nodes: dict[str, dict[str, Any]] = {}
    edges: dict[tuple[str, str, str], dict[str, Any]] = {}

    per_sheet_groups: dict[str, list[FormulaGroup]] = defaultdict(list)
    for grp in groups.values():
        per_sheet_groups[grp.sheet].append(grp)

    kept_groups: list[tuple[str, FormulaGroup]] = []
    for sheet, sheet_groups in per_sheet_groups.items():
        sheet_groups.sort(key=lambda g: (-len(g.cells), g.rep))
        kept = sheet_groups[:MAX_NODES_PER_SHEET]
        dropped = sheet_groups[MAX_NODES_PER_SHEET:]
        for grp in kept:
            rep_r, rep_c = grp.rep
            if len(grp.cells) == 1:
                node_id = f"c:{sheet}!{a1(rep_r, rep_c)}"
            else:
                node_id = f"g:{sheet}!{a1(rep_r, rep_c)}#{len(grp.cells)}"
            kept_groups.append((node_id, grp))
            for cell in grp.cells:
                cell_owner[sheet][cell] = node_id
        if dropped:
            n_cells = sum(len(g.cells) for g in dropped)
            misc_id = f"misc:{sheet}"
            nodes[misc_id] = {
                "id": misc_id,
                "kind": "misc",
                "sheet": sheet,
                "label": f"{len(dropped)} other patterns ({n_cells} cells)",
                "count": n_cells,
                "patterns": len(dropped),
            }
            warnings.append(
                f"Sheet '{sheet}': {len(dropped)} formula patterns aggregated "
                f"into a 'misc' node (limit {MAX_NODES_PER_SHEET})"
            )
            for grp in dropped:
                for cell in grp.cells:
                    cell_owner[sheet][cell] = misc_id

    ast_cache: dict[str, Any] = {}
    input_nodes: dict[str, str] = {}  # full A1 key -> node id

    def ensure_input_node(rect: Rect, opaque_label: str | None = None) -> str:
        label = opaque_label or rect.to_a1()
        node_id = input_nodes.get(label)
        if node_id:
            return node_id
        if opaque_label is not None:
            node_id = f"x:{opaque_label}"
            nodes[node_id] = {
                "id": node_id,
                "kind": "opaque",
                "label": opaque_label,
                "sheet": None,
            }
        else:
            node_id = f"i:{label}"
            node: dict[str, Any] = {
                "id": node_id,
                "kind": "input",
                "label": label,
                "sheet": rect.sheet,
                "addr": label.split("!")[-1],
                "count": rect.ncells,
                "values": _sample_range_values(resolver, rect),
            }
            if rect.ncells == 1 and rect.sheet is not None:
                node.update(resolver.describe(rect.sheet, rect.r1, rect.c1))
                _enrich_with_table(node, table_index, rect.sheet, rect.r1, rect.c1)
            nodes[node_id] = node
        input_nodes[label] = node_id
        return node_id

    def add_edge(src: str, dst: str, kind: str, approx: bool = False) -> None:
        if src == dst:
            return
        key = (src, dst, kind)
        e = edges.get(key)
        if e is None:
            edges[key] = {
                "id": f"e{len(edges)}",
                "source": src,
                "target": dst,
                "kind": kind,
                "approx": approx,
            }
        elif not approx:
            e["approx"] = False

    def resolve_rect_edges(rect: Rect, target_id: str, kind: str = "dep") -> None:
        """Create precedent → target edges for a referenced range."""
        sheet = rect.sheet
        if sheet not in sheet_dims:
            ensure_input_node(rect, opaque_label=rect.to_a1())
            add_edge(input_nodes[rect.to_a1()], target_id, kind)
            return
        clipped = rect.clipped(*sheet_dims[sheet])
        if clipped is None:
            return
        owners = cell_owner.get(sheet, {})
        if clipped.ncells <= SMALL_RANGE_CELLS:
            seen: set[str] = set()
            has_plain = False
            for r in range(clipped.r1, clipped.r2 + 1):
                for c in range(clipped.c1, clipped.c2 + 1):
                    owner = owners.get((r, c))
                    if owner is None:
                        has_plain = True
                    elif owner not in seen:
                        seen.add(owner)
                        add_edge(owner, target_id, kind)
            if has_plain:
                add_edge(ensure_input_node(clipped), target_id, kind)
        else:
            # Huge range: approximate intersection with node bounding boxes.
            for node_id, grp in kept_groups:
                if grp.sheet != sheet:
                    continue
                r1, c1, r2, c2 = grp.bbox
                if clipped.intersects(Rect(sheet, r1, c1, r2, c2)):
                    add_edge(node_id, target_id, kind, approx=True)
            add_edge(ensure_input_node(clipped), target_id, kind, approx=True)

    # defined names -----------------------------------------------------------
    name_nodes: dict[str, str] = {}
    for name, targets in defined_names.items():
        node_id = f"n:{name}"
        name_nodes[name.upper()] = node_id
        value_fields: dict[str, Any] = {"value": None}
        if targets:
            first = targets[0]
            if (
                first.sheet is not None
                and first.r1 == first.r2
                and first.c1 == first.c2
            ):
                value_fields = resolver.describe(first.sheet, first.r1, first.c1)
            else:
                val_samples = _sample_range_values(resolver, first)
                if val_samples:
                    value_fields = {"value": val_samples[0]["value"]}
        nodes[node_id] = {
            "id": node_id,
            "kind": "name",
            "label": name,
            "sheet": targets[0].sheet if targets else None,
            "targets": [t.to_a1() for t in targets],
            **value_fields,
        }
        for rect in targets:
            resolve_rect_edges(rect, node_id, kind="name")

    # formula nodes + edges -------------------------------------------------
    for node_id, grp in kept_groups:
        rep_r, rep_c = grp.rep
        formula = grp.formulas.get((rep_r, rep_c)) or next(iter(grp.formulas.values()))
        sheet = grp.sheet
        is_group = len(grp.cells) > 1
        try:
            ast = ast_cache.get(formula)
            if ast is None:
                ast = ast_cache[formula] = fz.parse(
                    formula if formula.startswith("=") else "=" + formula
                )
            ast_dict = ast.to_dict()
        except Exception:
            ast, ast_dict = None, None

        refs = _collect_ref_strings(ast_dict) if ast_dict else []
        rmin, cmin, rmax, cmax = grp.bbox
        agg_rects: list[Rect] = []
        for ref in refs:
            detail = parse_ref_detailed(ref, default_sheet=sheet)
            if detail is None:
                up = ref.upper()
                if up in name_nodes:
                    add_edge(name_nodes[up], node_id, "name")
                else:
                    opaque_id = ensure_input_node(
                        Rect(None, 1, 1, 1, 1), opaque_label=ref
                    )
                    add_edge(opaque_id, node_id, "dep")
                continue
            rect = (
                stretch_ref(detail, rep_r, rep_c, (rmin, rmax), (cmin, cmax))
                if is_group
                else detail.rect
            )
            agg_rects.append(rect)

        for rect in _merge_rects(agg_rects):
            resolve_rect_edges(rect, node_id)

        value_fields = resolver.describe(sheet, rep_r, rep_c, formula)
        samples = None
        if is_group:
            samples = []
            for r, c in itertools.islice(sorted(grp.cells), 3):
                samples.append(
                    {
                        "addr": a1(r, c),
                        **resolver.describe(sheet, r, c, grp.formulas.get((r, c))),
                    }
                )

        steps = None
        if ast_dict is not None:
            step_exprs = _collect_step_exprs(ast_dict)
            if step_exprs:
                resolver.preload_steps(step_exprs, sheet)
            steps = _decompose(ast_dict, sheet, resolver, defined_names)

        node: dict[str, Any] = {
            "id": node_id,
            "kind": "group" if is_group else "cell",
            "sheet": sheet,
            "addr": a1(rep_r, rep_c),
            "label": (
                f"{sheet}!{a1(rep_r, rep_c)}"
                + (f" x{len(grp.cells)}" if is_group else "")
            ),
            "formula": formula if formula.startswith("=") else "=" + formula,
            "r1c1": grp.r1c1,
            "count": len(grp.cells),
            "bbox": _bbox_a1(grp),
            **value_fields,
            "samples": samples,
            "steps": steps,
        }
        _enrich_with_table(node, table_index, sheet, rep_r, rep_c)
        nodes[node_id] = node

    # --- 6. VBA --------------------------------------------------------------
    vba_modules = extract_vba_modules(data, filename, warnings)
    vba_procs: list[VbaProc] = analyze_vba(vba_modules) if vba_modules else []
    # Node ids keep the declared spelling, but both lookups are keyed on the
    # lowercased name: VBA is case-insensitive, so Module1.Taux and
    # module1.TAUX designate the same procedure. proc_ids resolves a qualified
    # name, procs_by_name the unqualified ones _find_calls reports.
    proc_ids: dict[str, str] = {}
    procs_by_name: dict[str, list[str]] = defaultdict(list)
    for proc in vba_procs:
        qualified = f"{proc.module}.{proc.name}"
        pid = f"vp:{qualified}"
        proc_ids[qualified.lower()] = pid
        procs_by_name[proc.name.lower()].append(qualified.lower())
        nodes[pid] = {
            "id": pid,
            "kind": "vba",
            "label": f"{proc.module}.{proc.name}",
            "sheet": None,
            "module": proc.module,
            "proc": proc.name,
            "procKind": proc.kind,
            "lines": [proc.line_start, proc.line_end],
            "code": proc.code[:MAX_VBA_CODE_CHARS],
        }
    for proc in vba_procs:
        pid = proc_ids[f"{proc.module}.{proc.name}".lower()]
        for callee in proc.calls:
            target = _resolve_call(callee, proc.module, proc_ids, procs_by_name)
            if target is not None:
                add_edge(pid, target, "call")
        for ref in proc.refs:
            detail = parse_ref_detailed(ref.ref, default_sheet=ref.sheet)
            if detail is None or detail.rect.sheet is None:
                opaque_id = ensure_input_node(
                    Rect(None, 1, 1, 1, 1),
                    opaque_label=f"VBA:{ref.sheet or '?'}!{ref.ref}",
                )
                if ref.access == "write":
                    add_edge(pid, opaque_id, "vba-write")
                else:
                    add_edge(opaque_id, pid, "vba-read")
                continue
            if ref.access == "write":
                _resolve_vba_write(
                    detail.rect,
                    pid,
                    sheet_dims,
                    cell_owner,
                    add_edge,
                    ensure_input_node,
                )
            else:
                resolve_rect_edges(detail.rect, pid, kind="vba-read")

    if not engine_alive and resolver.n_recovered + resolver.n_unrecovered:
        warnings.append(
            f"Values recovered cell by cell: {resolver.n_recovered} recomputed, "
            f"{resolver.n_unrecovered} left to the value stored in the file"
        )
    uncomputed = resolver.uncomputed_warning()
    if uncomputed:
        warnings.append(uncomputed)

    graph = {
        "meta": {
            "filename": filename,
            "analyzedAt": datetime.datetime.now(datetime.UTC).isoformat(),
            "engine": "formualizer (Rust)",
            "warnings": warnings,
            "stats": {
                "sheets": sheet_stats,
                "totalFormulas": formula_count,
                "totalNodes": len(nodes),
                "totalEdges": len(edges),
                "groupedPatterns": sum(1 for _, g in kept_groups if len(g.cells) > 1),
                "vbaModules": len(vba_modules),
                "vbaProcs": len(vba_procs),
                "definedNames": len(defined_names),
                "tables": sum(len(t) for t in table_index.values()),
            },
        },
        "sheets": list(sheet_dims.keys()),
        "nodes": list(nodes.values()),
        "edges": list(edges.values()),
    }
    _v("nodes+edges+graph", _t)
    if verbose:
        print(
            f"[linexcel] total: {time.perf_counter() - _t0:.1f}s | "
            f"{len(nodes)} nodes | {len(edges)} edges | "
            f"{formula_count:,} formulas",
            file=sys.stderr,
        )
    return {"graph": graph, "engine": engine, "analysisId": uuid.uuid4().hex[:16]}


# ---------------------------------------------------------------------------
# Composite function decomposition
# ---------------------------------------------------------------------------

_STEP_KINDS = {"Function", "BinaryOp", "UnaryOp"}


def _collect_step_exprs(ast_dict: dict) -> list[str]:
    """First pass: collect every step expression ``_decompose`` will evaluate.

    The counter and ``_STEP_KINDS`` filter mirror ``_decompose`` exactly so
    the expressions batch-evaluated here are the same ones looked up later
    in ``_eval_raw``.  Order is pre-order (parent before children) — it does
    not match the post-order evaluation in ``_decompose`` but that is fine:
    the batch evaluates all at once and the cache is order-independent.
    """
    exprs: list[str] = []
    counter = itertools.count()

    def walk(node: dict) -> None:
        ntype = node.get("node_type")
        if ntype not in _STEP_KINDS:
            return
        if next(counter) >= MAX_STEPS_PER_FORMULA:
            return
        exprs.append(_render_expr(node))
        if ntype == "Function":
            children = node.get("args", [])
        elif ntype == "BinaryOp":
            children = [node.get("left"), node.get("right")]
        else:
            children = [node.get("operand") or node.get("expr")]
        for c in children:
            if c:
                walk(c)

    walk(ast_dict)
    return exprs


def _decompose(
    ast_dict: dict,
    sheet: str,
    resolver: _ValueResolver,
    defined_names: dict[str, list[Rect]] | None = None,
) -> dict | None:
    """Step tree: each function / operator becomes an evaluated step."""
    counter = itertools.count()

    def expr_of(node: dict) -> str:
        return _render_expr(node)

    def walk(node: dict, depth: int) -> dict | None:
        ntype = node.get("node_type")
        if ntype not in _STEP_KINDS:
            return None
        if next(counter) >= MAX_STEPS_PER_FORMULA:
            return None
        expr = expr_of(node)
        if ntype == "Function":
            label = node.get("name", "?")
            children_ast = node.get("args", [])
        elif ntype == "BinaryOp":
            label = node.get("operator", "?")
            children_ast = [node.get("left"), node.get("right")]
        else:
            label = node.get("operator", "?")
            children_ast = [node.get("operand") or node.get("expr")]
        children_ast = [c for c in children_ast if c]

        inputs = []
        children = []
        for child in children_ast:
            sub = walk(child, depth + 1)
            if sub is not None:
                children.append(sub)
            else:
                ctype = child.get("node_type")
                if ctype == "Reference":
                    ref = child.get("reference", "?")
                    preview, date_text = _ref_preview(
                        resolver, ref, sheet, defined_names
                    )
                    entry: dict[str, Any] = {"ref": ref, "value": preview}
                    if date_text is not None:
                        entry["date"] = date_text
                    inputs.append(entry)
                elif ctype == "Literal":
                    inputs.append({"literal": child.get("value")})

        value, evaluated = resolver.eval_expr(expr, sheet)
        return {
            "kind": ntype,
            "label": label,
            "expr": expr,
            "value": value,
            "evaluated": evaluated,
            "inputs": inputs,
            "children": children,
        }

    return walk(ast_dict, 0)


def _render_expr(node: dict) -> str:
    """Reconstruct the expression of an AST subtree (readable form)."""
    ntype = node.get("node_type")
    if ntype == "Function":
        args = ", ".join(_render_expr(a) for a in node.get("args", []))
        return f"{node.get('name', '?')}({args})"
    if ntype == "BinaryOp":
        return (
            f"{_render_expr(node.get('left', {}))} {node.get('operator', '?')} "
            f"{_render_expr(node.get('right', {}))}"
        )
    if ntype == "UnaryOp":
        operand = node.get("operand") or node.get("expr") or {}
        return f"{node.get('operator', '?')}{_render_expr(operand)}"
    if ntype == "Reference":
        return str(node.get("reference", "?"))
    if ntype == "Literal":
        v = node.get("value")
        if isinstance(v, str):
            return '"' + v.replace('"', '""') + '"'
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    if ntype == "Array":
        return "{...}"
    if ntype == "Paren":
        inner = node.get("expr") or node.get("inner") or {}
        return f"({_render_expr(inner)})"
    return "?"


def _scratch_eval(engine, expr: str, sheet: str) -> tuple[Any, bool]:
    """Evaluate an expression on its own in the scratch sheet.

    The cell is primed with a sentinel first: when the engine cannot compute
    an expression it leaves the previous value in place instead of reporting
    an error, and an unchanged sentinel is the only way to tell. The very
    first evaluation on a workbook holding a broken reference raises — the
    engine walks the whole dirty graph — while later ones stay isolated,
    hence the single retry.
    """
    try:
        qualified = qualify_sheet(expr, sheet)
    except Exception:
        return None, False
    for _ in range(2):
        try:
            engine.set_formula(SCRATCH_SHEET, 1, 1, f'="{SCRATCH_SENTINEL}"')
            if engine.evaluate_cell(SCRATCH_SHEET, 1, 1) != SCRATCH_SENTINEL:
                continue
            engine.set_formula(SCRATCH_SHEET, 1, 1, qualified)
            value = engine.evaluate_cell(SCRATCH_SHEET, 1, 1)
        except Exception:
            continue
        if value == SCRATCH_SENTINEL:
            return None, False
        return value, True
    return None, False


def _guard_fallback_expr(expr: str) -> str | None:
    """Fallback branch of a top-level IFERROR/IFNA, as Excel would show it."""
    # ponytail: only a top-level guard is recovered. With a nested one —
    # =IFERROR(SUM(IFERROR(NOSHEET!A1,0)),1) — the whole expression fails to
    # evaluate, so the outer fallback branch is taken (1) where Excel lets the
    # inner guard absorb the broken reference and returns 0. That gap is the
    # accepted ceiling of this recovery.
    try:
        ast_dict = fz.parse(expr).to_dict()
    except Exception:
        return None
    if not isinstance(ast_dict, dict) or ast_dict.get("node_type") != "Function":
        return None
    if str(ast_dict.get("name", "")).upper() not in GUARD_FUNCTIONS:
        return None
    args = ast_dict.get("args") or []
    if len(args) < 2:
        return None
    return "=" + _render_expr(args[1])


#: A bracketed group: an external workbook (``[Budget.xlsx]Sheet1!A1``) or a
#: structured table reference (``SalesTable[Revenue]``). Neither resolves to a
#: cell the engine holds.
_BRACKETED_RE = re.compile(r"\[[^\]]*\]")
#: A 3-D sheet span — ``'First:Last'!A1`` or ``First:Last!A1``. The engine reads
#: the span as one sheet name and reports it missing.
_SPAN_RE = re.compile(r"'[^']*:[^']*'!|(?<![\w$)])[A-Za-z_][\w.]*:[A-Za-z_][\w.]*!")
#: Error guards. A formula using one may well have a defined value despite a
#: reference that cannot be resolved, so it is never isolated.
_GUARD_RE = re.compile(r"\b(?:IFERROR|IFNA|ISERROR|ISERR|ISNA)\s*\(", re.IGNORECASE)
#: A sheet qualifier, quoted or bare. The lookbehind keeps ``#REF!`` out of the
#: bare form: that is an error literal, not a sheet called REF.
_SHEET_QUALIFIER_RE = re.compile(r"'((?:[^']|'')+)'!|(?<![#\w.$])([A-Za-z_][\w.]*)!")


def _quarantine_unresolvable(
    engine, sheet_dims: dict[str, tuple[int, int]], engine_sheets: set[str]
) -> dict[tuple[str, int, int], str]:
    """Blank the formulas that stop ``evaluate_all``, returning what was removed.

    Only ever called after a global evaluation has already failed, so both ways
    of being wrong are safe: quarantining a formula that would in fact have
    evaluated costs it the same per-cell recovery every cell was getting anyway,
    and missing one simply leaves the retry failing as before.

    The formula text is returned rather than restored — ``set_formula`` does not
    put it back once the cell holds a value — and the scan re-injects it so the
    graph still shows what the cell really contains.
    """
    suspects: dict[tuple[str, int, int], str] = {}
    for sheet, (max_row, max_col) in sheet_dims.items():
        if sheet not in engine_sheets:
            continue
        try:
            fsheet = engine.sheet(sheet)
        except Exception:
            continue
        for r0 in range(1, max_row + 1, SCAN_CHUNK_ROWS):
            r1 = min(r0 + SCAN_CHUNK_ROWS - 1, max_row)
            try:
                rows = fsheet.get_formulas(fz.RangeAddress(sheet, r0, 1, r1, max_col))
            except Exception:
                break
            for i, row_vals in enumerate(rows):
                for j, formula in enumerate(row_vals):
                    if formula and _is_unresolvable(formula, engine_sheets):
                        suspects[(sheet, r0 + i, j + 1)] = formula
    for (sheet, row, col), _formula in suspects.items():
        try:
            engine.set_value(sheet, row, col, None)
        except Exception:
            continue
    return suspects


def _is_unresolvable(formula: str, engine_sheets: set[str]) -> bool:
    """Whether ``formula`` names something the engine cannot resolve.

    A guarded formula is never isolated, however broken its reference looks.
    ``IFERROR(NOSHEET!A1, 456)`` *has* a correct value — 456 — and blanking it
    does not merely cost that cell its own value: every range that spans it
    silently loses a term, so a `SUM` over the column returns a different
    number. Quarantine is only safe where there was no value to lose.
    """
    if _GUARD_RE.search(formula):
        return False
    if _BRACKETED_RE.search(formula) or _SPAN_RE.search(formula):
        return True
    # A sheet qualifier naming a sheet the engine never loaded.
    for match in _SHEET_QUALIFIER_RE.finditer(formula):
        name = match.group(1)
        name = name.replace("''", "'") if name else match.group(2)
        if name and name not in engine_sheets:
            return True
    return False


def _ensure_scratch(engine) -> bool:
    try:
        engine.add_sheet(SCRATCH_SHEET)
        return True
    except Exception:
        return SCRATCH_SHEET in set(engine.sheet_names)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_table_index(
    data: bytes,
) -> dict[str, list[dict[str, Any]]]:
    """Per-sheet table list, for cell→table enrichment of lineage nodes.

    openpyxl read-only mode drops ``ws.tables``, so the workbook is opened
    once in normal mode. Each table dict carries the bounds and headers that
    :func:`linexcel.insights.detect_tables` extracts.
    """
    from linexcel.insights import detect_tables

    index: dict[str, list[dict[str, Any]]] = {}
    try:
        wb = load_workbook(io.BytesIO(data), read_only=False, data_only=False)
    except Exception:
        return index
    try:
        for ws in wb.worksheets:
            try:
                tables = detect_tables(ws)
            except Exception:
                continue
            if tables:
                index[ws.title] = tables
    finally:
        wb.close()
    return index


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


def _collect_ref_strings(ast_dict: dict) -> list[str]:
    """Every reference a formula makes, excluding names it binds itself.

    The parser reports a ``LET`` binding — and a ``LAMBDA`` parameter — as a
    ``Reference`` node, because syntactically it looks like one. It is not: the
    name is local to the formula and points at no cell, so treating it as a
    reference put a node on the graph for every intermediate a modeller happened
    to name. A formula's own bindings are its business; the graph shows the
    cells, and the LET is visible in that cell's step decomposition.
    """
    bound = _bound_names(ast_dict)
    refs: list[str] = []

    def walk(node) -> None:
        if isinstance(node, dict):
            if node.get("node_type") == "Reference":
                ref = node.get("reference")
                if ref and str(ref).upper() not in bound:
                    refs.append(str(ref))
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    walk(ast_dict)
    # dedupe preserving order
    seen: set[str] = set()
    out = []
    for r in refs:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out


def _bound_names(ast_dict: dict) -> set[str]:
    """Upper-cased names bound by any ``LET`` or ``LAMBDA`` in the formula.

    ``LET(n1, v1, [n2, v2, ...], calc)`` binds the even-indexed arguments before
    the trailing calculation; ``LAMBDA(p1, ..., calc)`` binds every argument but
    the last. Only bare identifiers count — ``LET(A1, ...)`` is not legal Excel,
    so a name that parses as a real reference is left alone.
    """
    bound: set[str] = set()

    def walk(node) -> None:
        if isinstance(node, dict):
            if node.get("node_type") == "Function":
                name = str(node.get("name") or "").upper()
                args = node.get("args") or []
                if name == "LET":
                    indices = range(0, max(len(args) - 1, 0), 2)
                elif name == "LAMBDA":
                    indices = range(max(len(args) - 1, 0))
                else:
                    indices = range(0)
                for index in indices:
                    arg = args[index]
                    if not isinstance(arg, dict):
                        continue
                    if arg.get("node_type") != "Reference":
                        continue
                    text = str(arg.get("reference") or "")
                    if text and parse_ref_detailed(text, default_sheet="") is None:
                        bound.add(text.upper())
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    walk(ast_dict)
    return bound


def _merge_rects(rects: list[Rect]) -> list[Rect]:
    seen: set[tuple] = set()
    out = []
    for r in rects:
        key = (r.sheet, r.r1, r.c1, r.r2, r.c2)
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def _bbox_a1(grp: FormulaGroup) -> str:
    r1, c1, r2, c2 = grp.bbox
    if (r1, c1) == (r2, c2):
        return a1(r1, c1)
    return f"{a1(r1, c1)}:{a1(r2, c2)}"


def _sample_range_values(resolver: _ValueResolver, rect: Rect) -> list:
    if rect.sheet is None:
        return []
    out = []
    for r in range(rect.r1, min(rect.r2, rect.r1 + MAX_VALUE_SAMPLE - 1) + 1):
        for c in range(rect.c1, min(rect.c2, rect.c1 + MAX_VALUE_SAMPLE - 1) + 1):
            if len(out) >= MAX_VALUE_SAMPLE:
                return out
            value, source, date_text = resolver.value(rect.sheet, r, c)
            out.append(
                {
                    "addr": a1(r, c),
                    "value": value,
                    "source": source,
                    "date": date_text,
                }
            )
    return out


def _ref_preview(
    resolver: _ValueResolver,
    ref: str,
    sheet: str,
    defined_names: dict[str, list[Rect]] | None = None,
) -> tuple[Any, str | None]:
    """Preview of a referenced cell or range: ``(value, date_text)``."""
    detail = parse_ref_detailed(ref, default_sheet=sheet)
    if detail is None:
        # may be a defined name: show the value of its target
        if defined_names:
            for name, rects in defined_names.items():
                if name.upper() == ref.upper() and rects:
                    rect = rects[0]
                    if rect.ncells == 1:
                        value, _, date_text = resolver.value(
                            rect.sheet or sheet, rect.r1, rect.c1
                        )
                        return value, date_text
                    return {"range": rect.to_a1(), "n": rect.ncells}, None
        return None, None
    rect = detail.rect
    if rect.ncells == 1:
        value, _, date_text = resolver.value(rect.sheet or sheet, rect.r1, rect.c1)
        return value, date_text
    return {"range": rect.to_a1(), "n": rect.ncells}, None


def _resolve_call(
    callee: str,
    caller_module: str,
    proc_ids: dict[str, str],
    procs_by_name: dict[str, list[str]],
) -> str | None:
    """Resolve a VBA call to a procedure node id.

    A callee already qualified with its module (``Module1.Taux``) resolves
    directly. VBA looks an unqualified name up in the calling module first,
    then in the other modules; a name declared by several other modules stays
    unresolved rather than pointing at an arbitrary one. All lookups are
    case-insensitive, as the language is.
    """
    if "." in callee:
        return proc_ids.get(callee.lower())
    same_module = f"{caller_module}.{callee}".lower()
    if same_module in proc_ids:
        return proc_ids[same_module]
    candidates = procs_by_name.get(callee.lower(), [])
    if len(candidates) == 1:
        return proc_ids[candidates[0]]
    return None


def _resolve_vba_write(
    rect: Rect, pid: str, sheet_dims, cell_owner, add_edge, ensure_input_node
) -> None:
    """A VBA write feeds the target cells: edge proc → target."""
    sheet = rect.sheet
    if sheet not in sheet_dims:
        opaque = ensure_input_node(rect, opaque_label=rect.to_a1())
        add_edge(pid, opaque, "vba-write")
        return
    clipped = rect.clipped(*sheet_dims[sheet]) or rect
    owners = cell_owner.get(sheet, {})
    seen: set[str] = set()
    has_plain = False
    if clipped.ncells <= SMALL_RANGE_CELLS:
        for r in range(clipped.r1, clipped.r2 + 1):
            for c in range(clipped.c1, clipped.c2 + 1):
                owner = owners.get((r, c))
                if owner is None:
                    has_plain = True
                elif owner not in seen:
                    seen.add(owner)
                    add_edge(pid, owner, "vba-write")
    else:
        has_plain = True
    if has_plain:
        add_edge(pid, ensure_input_node(clipped), "vba-write")


def serial_to_date_text(serial: Any, epoch_1904: bool = False) -> str | None:
    """Excel serial number → ``YYYY-MM-DD``, or None if it is not a date."""
    if isinstance(serial, bool) or not isinstance(serial, (int, float)):
        return None
    days = float(serial)
    if days != days or days in (float("inf"), float("-inf")):
        return None
    if epoch_1904:
        if days < 0:
            return None
        base = EXCEL_EPOCH_1904
    else:
        if days < 1 or days == 60:
            return None
        base = EPOCH_EARLY_1900 if days < 60 else EXCEL_EPOCH_1900
    try:
        return (base + datetime.timedelta(days=days)).date().isoformat()
    except (OverflowError, ValueError):
        return None


def _is_date_format(number_format: Any) -> bool:
    if not number_format:
        return False
    try:
        return bool(is_date_format(number_format))
    except Exception:
        return False


def _date_text_of(value: Any) -> str | None:
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    return None


def _error_kind(value: Any) -> str | None:
    """Kind of an engine error value, or ``None`` when it is not one.

    The engine reports errors as a plain ``{"type": "Error", "kind": ...}``
    dict, which stringifies to Python source rather than to anything a reader
    of a spreadsheet recognises — hence every path out of the engine going
    through here first.
    """
    if isinstance(value, dict) and value.get("type") == "Error":
        kind = value.get("kind")
        return kind if isinstance(kind, str) else ""
    return None


def _excel_error_text(value: Any) -> str | None:
    """Spreadsheet text of an engine error value that a cell can hold."""
    kind = _error_kind(value)
    return None if kind is None else ERROR_KIND_TEXT.get(kind)


def _is_uncomputed(value: Any) -> bool:
    """True when the engine reported a limitation instead of a result.

    An unknown kind counts too: a kind this table does not know is one whose
    spreadsheet meaning we cannot vouch for, and inventing a value for it is
    worse than admitting the cell was not recalculated.
    """
    kind = _error_kind(value)
    if kind is None:
        return False
    return kind in UNCOMPUTED_ERROR_KINDS or kind not in ERROR_KIND_TEXT


def _values_differ(raw: Any, cached: Any, date_text: str | None) -> bool:
    """True when a recalculated value contradicts the one stored in the file."""
    cached_date = _date_text_of(cached)
    if cached_date is not None:
        return date_text is not None and date_text != cached_date
    # An error is a value a cell genuinely holds, and openpyxl reads the stored
    # one back as the same text Excel shows — so the two are directly
    # comparable, and a recalculated #DIV/0! over a stored #DIV/0! is agreement,
    # not a disagreement nobody can explain.
    error_text = _excel_error_text(raw)
    if error_text is not None or isinstance(cached, str) and cached in EXCEL_ERRORS:
        return error_text != cached
    if isinstance(raw, bool) or isinstance(cached, bool):
        return False
    if isinstance(raw, (int, float)) and isinstance(cached, (int, float)):
        return abs(float(raw) - float(cached)) > 1e-9
    return False


def _fmt_value(value: Any) -> str:
    return _date_text_of(value) or _excel_error_text(value) or str(value)


def _jsonable(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        is_nan_or_inf = value != value or value in (float("inf"), float("-inf"))
        if isinstance(value, float) and is_nan_or_inf:
            return str(value)
        return value
    error_text = _excel_error_text(value)
    if error_text is not None:
        return error_text
    return str(value)
