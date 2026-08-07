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
# Serials below 61 sit before Excel's phantom 1900-02-29; no real date matches.
MIN_DATE_SERIAL_1900 = 61
GUARD_FUNCTIONS = {"IFERROR", "IFNA"}


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
    """Read the file's cached values once, keyed by (sheet, row, col)."""
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
        self.n_recovered = 0
        self.n_unrecovered = 0

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
        # dates stay dates: every serializer of the graph uses ``default=str``
        if isinstance(raw, (datetime.datetime, datetime.date)):
            return raw
        return _jsonable(raw)

    def eval_expr(self, expr: str, sheet: str) -> tuple[Any, bool]:
        """Evaluate an expression in the scratch sheet, if budget allows."""
        raw, ok = self._eval_raw(expr, sheet)
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
        return _scratch_eval(self.engine, expr, sheet)

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
        if self._engine_alive:
            try:
                raw = self.engine.evaluate_cell(sheet, row, col)
                if raw is not None:
                    return raw, "engine"
            except Exception:
                # The first failure poisons the engine for good: every later
                # whole-graph evaluation would raise the same way.
                self._engine_alive = False
        expr = formula if formula.startswith("=") else "=" + formula
        return self._remember(sheet, row, col, self._eval_formula(sheet, expr, 0))

    def _eval_formula(
        self, sheet: str, expr: str, depth: int
    ) -> tuple[Any, str | None]:
        """Evaluate one formula on its own, precedents resolved first."""
        if not self._engine_alive:
            self._resolve_precedents(sheet, expr, depth)
        raw, ok = self._eval_raw(expr, sheet)
        if ok and raw is not None:
            return raw, "engine"
        fallback = _guard_fallback_expr(expr)
        if fallback is not None:
            raw, ok = self._eval_raw(fallback, sheet)
            if ok and raw is not None:
                return raw, "fallback"
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
        if raw is None:
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


def a1(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def analyze_workbook(data: bytes, filename: str = "workbook.xlsx") -> dict[str, Any]:
    """Full analysis: returns the JSON-serializable graph and the engine."""
    warnings: list[str] = []

    # --- 1. structure -----------------------------------------------------
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

    # values the file itself carries: last resort, and the only source of
    # dates and of what the user actually saw on screen
    cached = load_cached_values(data)

    # --- 2. computation engine -------------------------------------------
    engine = fz.Workbook.from_bytes(data)
    engine_sheets = set(engine.sheet_names)
    engine_alive = True
    try:
        engine.evaluate_all()
    except Exception as exc:  # graph remains useful without values
        warnings.append(f"Global evaluation incomplete: {exc}")
        # A failed global evaluation does not just drop the values: the engine
        # then reports no formula at all, which would leave the graph empty.
        # Rebuilding from the bytes gives the formulas back; values are
        # recovered cell by cell further down.
        engine_alive = False
        engine = fz.Workbook.from_bytes(data)

    scratch_ready = _ensure_scratch(engine)
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
        nodes[node_id] = node

    # --- 6. VBA --------------------------------------------------------------
    vba_modules = extract_vba_modules(data, filename)
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
            },
        },
        "sheets": list(sheet_dims.keys()),
        "nodes": list(nodes.values()),
        "edges": list(edges.values()),
    }
    return {"graph": graph, "engine": engine, "analysisId": uuid.uuid4().hex[:16]}


# ---------------------------------------------------------------------------
# Composite function decomposition
# ---------------------------------------------------------------------------

_STEP_KINDS = {"Function", "BinaryOp", "UnaryOp"}


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


def _ensure_scratch(engine) -> bool:
    try:
        engine.add_sheet(SCRATCH_SHEET)
        return True
    except Exception:
        return SCRATCH_SHEET in set(engine.sheet_names)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


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
    out: dict[str, list[Rect]] = {}
    try:
        items = owb.defined_names.items()
    except Exception:
        return out
    for name, dn in items:
        if name.startswith("_xlnm."):
            continue
        rects: list[Rect] = []
        try:
            for sheet, coord in dn.destinations:
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


def _collect_ref_strings(ast_dict: dict) -> list[str]:
    refs: list[str] = []

    def walk(node) -> None:
        if isinstance(node, dict):
            if node.get("node_type") == "Reference":
                ref = node.get("reference")
                if ref:
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
        if days < MIN_DATE_SERIAL_1900:
            return None
        base = EXCEL_EPOCH_1900
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


def _values_differ(raw: Any, cached: Any, date_text: str | None) -> bool:
    """True when a recalculated value contradicts the one stored in the file."""
    cached_date = _date_text_of(cached)
    if cached_date is not None:
        return date_text is not None and date_text != cached_date
    if isinstance(raw, bool) or isinstance(cached, bool):
        return False
    if isinstance(raw, (int, float)) and isinstance(cached, (int, float)):
        return abs(float(raw) - float(cached)) > 1e-9
    return False


def _fmt_value(value: Any) -> str:
    return _date_text_of(value) or str(value)


def _jsonable(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        is_nan_or_inf = value != value or value in (float("inf"), float("-inf"))
        if isinstance(value, float) and is_nan_or_inf:
            return str(value)
        return value
    return str(value)
