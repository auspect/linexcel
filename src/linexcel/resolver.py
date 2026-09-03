"""The value of a cell, and where it came from — under budget.

Single entry point for reading a cell: engine value, per-cell recalculation,
isolated re-evaluation in the scratch sheet (with the IFERROR/IFNA fallback
branch), and finally the value cached in the file. Extracted mechanically
from analyzer.py.
"""

from __future__ import annotations

import datetime
import re
import time
from pathlib import Path
from typing import Any

import formualizer as fz

from linexcel.decompose import (
    SCRATCH_SENTINEL,
    SCRATCH_SHEET,
    _guard_fallback_expr,
    _scratch_eval,
)
from linexcel.external import (
    ExternalBook,
    ExternalRef,
    parse_external_refs,
    read_workbook_values,
)
from linexcel.loader import CachedValues
from linexcel.refs import a1, parse_ref, parse_ref_detailed
from linexcel.rewrite import qualify_sheet
from linexcel.values import (
    EXCEL_EPOCH_1900,
    EXCEL_ERRORS,
    _date_text_of,
    _fmt_value,
    _is_uncomputed,
    _jsonable,
    readings_agree,
    serial_to_date_text,
)

MAX_SCRATCH_EVALS = 4_000
#: Wall-clock ceiling on the step decomposition, in seconds. A count of
#: evaluations cannot bound time: each one asks the engine to walk the dirty
#: dependency graph, so on a workbook of cumulative sums — running totals, the
#: most ordinary thing a spreadsheet does — one evaluation costs O(graph) and
#: the count budget expires hours after anyone stopped waiting. This is the
#: bound that holds whatever the file looks like. Past it, cells keep their
#: values and lose only their step-by-step breakdown, and the report says so.
DEFAULT_STEP_SECONDS = 300.0
MAX_VALUE_WARNINGS = 25
# Chained recovery: how deep the precedent walk goes, and how wide a referenced
# range may be before its cells are left to the engine. Both only bound the
# work; a cell that is skipped simply keeps the value the engine reports.
MAX_CHAIN_DEPTH = 24
MAX_CHAIN_RANGE_CELLS = 4_096
#: Functions that answer with the clock or with chance. Excel calls a wider set
#: "volatile" — OFFSET, INDIRECT, CELL and INFO recalculate on every edit too —
#: but those return the same value for the same workbook, so linexcel keeps
#: computing them. These cannot be checked against anything: a `=TODAY()`
#: recomputed today never matches a file saved last week, and calling that a
#: divergence blames the workbook for the calendar.
VOLATILE_FUNCTIONS = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY"})
_VOLATILE_RE = re.compile(
    r"(?<![\w.])(?:_xlfn\.)?(?:" + "|".join(sorted(VOLATILE_FUNCTIONS)) + r")\s*\(",
    re.IGNORECASE,
)
#: A double-quoted Excel string, doubled quotes included. Blanked before the
#: volatile scan so `="TODAY() is volatile"` is not read as a call.
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
#: How many uncomputable cells the warning names before it stops listing them.
MAX_UNCOMPUTED_LISTED = 10


class _Budget:
    """What the step decomposition is allowed to spend, in calls and in time.

    The count came first and is not enough on its own: it bounds how many
    expressions are evaluated, not how long one takes, and a dense dependency
    graph makes each one arbitrarily expensive. Both are checked, and either
    running out ends the decomposition — not the analysis, which finishes with
    values intact and a warning naming what it stopped doing.
    """

    def __init__(self, limit: int, seconds: float | None = None) -> None:
        self.left = limit
        self.seconds = seconds
        self.deadline = None if seconds is None else time.monotonic() + seconds
        #: Which of the two ran out, for the warning to be able to say.
        self.spent: str | None = None

    def take(self, count: int = 1) -> bool:
        """Claim ``count`` evaluations, or refuse once the count is gone.

        Deliberately not time-bounded. This is the path that *recovers values*
        — the core of a report — and a clock that stopped it would trade the
        answer for the speed of getting it. Only the decomposition, which
        produces detail on top of an answer already found, honours the
        deadline; a test that asserted values survive a zero deadline is what
        caught the two being confused.
        """
        if self.left < count:
            self.spent = self.spent or "calls"
            return False
        self.left -= count
        return True

    @property
    def expired(self) -> bool:
        """True once the decomposition has spent its wall-clock allowance."""
        if self.deadline is None:
            return False
        if time.monotonic() > self.deadline:
            self.spent = "time"
            return True
        return False

    def warning(self) -> str | None:
        """What to tell someone whose report is missing its decompositions."""
        if self.spent is None:
            return None
        reason = (
            f"{self.seconds:.0f}s spent on it"
            if self.spent == "time"
            else f"{MAX_SCRATCH_EVALS:,} evaluations"
        )
        return (
            f"Step-by-step decomposition stopped after {reason}: the cells "
            f"past that point keep their values and show no breakdown. This "
            f"happens on workbooks whose formulas depend on each other in long "
            f"chains — running totals, for instance — where each step costs "
            f"the engine a walk of the whole graph. Raise it with "
            f"--time-budget, or read the report as it is."
        )


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
        externals: dict[str, ExternalBook] | None = None,
        refs_files: dict[str, Path] | None = None,
    ):
        self.engine = engine
        self.engine_sheets = engine_sheets
        self.cached = cached
        self.warnings = warnings
        self.budget = budget
        self.scratch_ready = scratch_ready
        self.sheet_dims = sheet_dims or {}
        self.externals = externals or {}
        self.refs_files = refs_files or {}
        self._declared_by_name = {
            book.name.lower(): book for book in self.externals.values() if book.name
        }
        self._named_books: dict[str, ExternalBook] = {}
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
        if formula is None:
            formula = self._formula_at(sheet, row, col)
        # A volatile formula answers differently every time it is computed, so
        # recomputing it says nothing about the workbook: `=TODAY()` recalculated
        # today cannot agree with a file saved last week, and reporting that as a
        # divergence blames the file for the calendar. The stored value is kept
        # and labelled as the only reading there is.
        if formula and _is_volatile(formula):
            value, _source, date_text = self._from_cache(sheet, row, col)
            return value, "volatile", date_text
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
            # Decided here rather than in the viewer's JavaScript: the same
            # question was being answered twice, with two different answers.
            fields["cachedAgreement"] = readings_agree(value, cached, date_text)
        if date_text is not None:
            fields["valueDate"] = date_text
        return fields

    def external_books(self, formula: str | None) -> list[dict[str, Any]]:
        """The other workbooks a formula reads, and how far each one was read.

        Named even when unreadable: "this cell depends on Budget FY26.xlsx,
        which linexcel was not given" is the answer a reader needs, and it is
        one a grey ``[1]`` node never gave.
        """
        if not formula:
            return []
        seen: dict[str, dict[str, Any]] = {}
        for ref in parse_external_refs(formula):
            book = self._book_for(ref)
            name = book.name if book else _external_name(ref)
            if name in seen:
                continue
            entry: dict[str, Any] = {"name": name}
            if book is not None and book.target:
                entry["path"] = book.target
            elif ref.directory:
                entry["path"] = ref.directory + name
            if book is not None and book.resolved:
                entry["read"] = "folder"
                entry["file"] = str(book.path)
            elif book is not None and book.cached:
                entry["read"] = "cache"
            else:
                entry["read"] = "none"
            seen[name] = entry
        return list(seen.values())

    def external_value(self, ref: ExternalRef) -> tuple[Any, str | None]:
        """``(value, source)`` for one external reference, or ``(None, None)``."""
        book = self._book_for(ref)
        position = _a1_position(ref.cell)
        if book is None or position is None:
            return None, None
        return book.value(ref.sheet, *position)

    def external_workbooks(self) -> list[ExternalBook]:
        """Every workbook seen, whether the file declared the link or not."""
        return list(self.externals.values()) + list(self._named_books.values())

    def substitute_externals(self, expr: str) -> str:
        """Replace external references by the values they resolve to.

        The engine cannot follow a link to another workbook — nothing in the
        file it loaded points there — so the reference is turned into the
        literal it stands for before the expression is evaluated. A reference
        that resolves to nothing is left alone: the evaluation then fails the
        way it did before, which is the honest outcome.
        """
        refs = parse_external_refs(expr)
        if not refs:
            return expr
        for ref in refs:
            book = self._book_for(ref)
            if book is None:
                continue
            position = _a1_position(ref.cell)
            if position is None:
                continue
            value, source = book.value(ref.sheet, *position)
            if source is None:
                continue
            expr = expr.replace(ref.text, _as_literal(value))
        return expr

    def _book_for(self, ref: ExternalRef) -> ExternalBook | None:
        """The workbook a reference points at, read on first use.

        A file that registered its links refers to them by index — ``[1]`` —
        and those are known up front. A file that did not spells the name out,
        ``[Budget FY26.xlsx]``, and then the reference folder is the only place
        to look; the workbook is read once and remembered.
        """
        declared = self.externals.get(ref.book) or self._declared_by_name.get(
            ref.book.lower()
        )
        if declared is not None:
            return declared
        if ref.book.isdigit():
            return None  # an index whose link part the file does not carry
        key = ref.book.lower()
        book = self._named_books.get(key)
        if book is None:
            book = ExternalBook(
                key=ref.book, target=ref.directory + ref.book, name=ref.book
            )
            path = self.refs_files.get(key)
            if path is not None:
                try:
                    book.values = read_workbook_values(path)
                    book.path = path
                except Exception as exc:
                    self.warnings.append(
                        f"External workbook '{ref.book}' could not be read: {exc}"
                    )
            self._named_books[key] = book
        return book

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

        This is the one door into the scratch sheet that the decomposition
        uses, and so the one place the wall-clock ceiling belongs. Gating the
        batch alone left the slow path running, which bounded the fast half of
        a runaway and none of the slow one.
        """
        if self.budget.expired:
            return None, False
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
        # Keyed on the expression as written, evaluated on the one the engine
        # can take: a link to another workbook becomes the value it stands for.
        return _scratch_eval(self.engine, self.substitute_externals(expr), sheet)

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
                qualified = qualify_sheet(self.substitute_externals(e), sheet)
            except Exception:
                continue
            try:
                # Primed with the sentinel first, exactly as ``_scratch_eval``
                # does: an expression the engine will not take — a structured
                # reference it cannot resolve, say — makes ``set_formula``
                # no-op rather than raise, and the cell then still holds the
                # step value *another node* left in that column. Reported as
                # this step's own result, that is a value from elsewhere.
                self.engine.set_formula(
                    SCRATCH_SHEET, 2, i + 1, f'="{SCRATCH_SENTINEL}"'
                )
                self.engine.set_formula(SCRATCH_SHEET, 2, i + 1, qualified)
            except Exception:
                continue
            targets.append((SCRATCH_SHEET, 2, i + 1))
            valid.append(e)
        if not targets:
            return
        # Claim the batch before running it. This call was invisible to the
        # budget: one evaluate_cells per node, each walking the dirty graph,
        # is exactly the work the budget exists to bound — and skipping it is
        # how a 4,000-evaluation ceiling turned into three hours.
        if self.budget.expired or not self.budget.take(len(targets)):
            return
        try:
            results = self.engine.evaluate_cells(targets)
        except Exception:
            return  # a broken reference poisons the batch — fall back
        for e, val in zip(valid, results):
            if val is not None and val != SCRATCH_SENTINEL and not _is_uncomputed(val):
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
        # A value that came out of another workbook is not the engine's own
        # reading of this file, and the card has to be able to say so.
        computed = "external" if parse_external_refs(expr) else "engine"
        raw, ok = self._eval_raw(expr, sheet)
        uncomputed = raw if ok and raw is not None and _is_uncomputed(raw) else None
        if ok and raw is not None and not _is_uncomputed(raw):
            return raw, computed
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
        if cached is None or readings_agree(raw, cached, date_text) != "differ":
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


def _external_warning(
    workbooks: list[ExternalBook], refs_dir: str | Path | None
) -> str | None:
    """One line naming every workbook this file reads, and how far each got.

    A workbook that depends on files nobody handed over is the common case in
    practice, and the values above such a link are only as good as the cache
    Excel left behind. Saying so is the point.
    """
    books = {b.name: b for b in workbooks if b.name}
    if not books:
        return None
    read = [b.name for b in books.values() if b.resolved]
    cached = [b.name for b in books.values() if not b.resolved and b.cached]
    missing = [b.name for b in books.values() if not b.resolved and not b.cached]
    parts = [f"This workbook reads {len(books)} external workbook(s)."]
    if read:
        parts.append(f"Read from the reference folder: {', '.join(sorted(read))}.")
    if cached:
        parts.append(
            f"Not read, values taken from the cache Excel left in the file: "
            f"{', '.join(sorted(cached))}."
        )
    if missing:
        parts.append(
            f"Neither read nor cached, so cells reading them have no value: "
            f"{', '.join(sorted(missing))}."
        )
        if refs_dir is None:
            parts.append("Pass refs_dir= (CLI: --refs-dir) to resolve them.")
    return " ".join(parts)


def _external_name(ref: ExternalRef) -> str:
    """The file name an external reference names, index or path alike."""
    return ref.book if not ref.book.isdigit() else f"[{ref.book}]"


def _a1_position(cell: str) -> tuple[int, int] | None:
    """``B4`` → ``(4, 2)``; a range takes its top-left cell."""
    rect = parse_ref(cell.split(":")[0])
    return (rect.r1, rect.c1) if rect is not None else None


def _as_literal(value: Any) -> str:
    """A value as the engine would read it back in a formula."""
    if value is None:
        return "0"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, datetime.datetime):
        return repr(_serial_of(value))
    text = str(value)
    if text in EXCEL_ERRORS:
        return text
    return '"' + text.replace('"', '""') + '"'


def _serial_of(moment: datetime.datetime) -> float:
    """A datetime as the 1900-epoch serial a formula computes with."""
    delta = moment - EXCEL_EPOCH_1900
    return delta.days + delta.seconds / 86_400


def _is_volatile(formula: str) -> bool:
    """Whether a formula answers with the clock or with chance.

    Text is blanked first: a cell spelling out ``="uses TODAY()"`` holds a
    string, not a call.
    """
    return bool(_VOLATILE_RE.search(_STRING_LITERAL_RE.sub('""', formula)))
