"""Boot the formualizer engine, and make it survive a broken reference.

Extracted mechanically from analyzer.py: instantiate the engine from the
workbook bytes, run the whole-workbook evaluation, and — when that fails
because one formula names something the engine cannot resolve — quarantine
the offending cells and retry, so the rest of the file keeps its computed
values.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import formualizer as fz

from linexcel.decompose import SCRATCH_SHEET

# Guards to stay responsive on large workbooks.
SCAN_CHUNK_ROWS = 20_000
#: Ceiling on one ``get_formulas`` call, in cells. The engine hands back a
#: dense grid of Python strings, so a 16,384-column sheet read 20,000 rows at
#: a time would materialize 327 million of them in one go.
SCAN_CHUNK_CELLS = 1_000_000

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


@dataclass
class EngineSession:
    engine: object
    engine_sheets: set[str]
    engine_alive: bool
    quarantined: dict[tuple[str, int, int], str]
    scratch_ready: bool


def boot_engine(
    data: bytes,
    sheet_dims: dict[str, tuple[int, int]],
    warnings: list[str],
) -> EngineSession:
    """Instantiate the engine and run its whole-workbook evaluation.

    ``evaluate_all`` is all-or-nothing, and it gives up on the *first*
    reference it cannot resolve — so a single formula pointing at another
    workbook costs every other cell in the file its computed value. When that
    happens, the offending cells are set aside and the pass retried, leaving
    only them to the slower per-cell recovery.
    """
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
    return EngineSession(
        engine, engine_sheets, engine_alive, quarantined, scratch_ready
    )


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


def _chunk_rows(max_col: int) -> int:
    """How many rows to read per ``get_formulas`` call on a sheet that wide.

    Rows alone are the wrong unit: the engine returns a dense grid of Python
    strings, so 20,000 rows of a 16,384-column sheet is 327 million of them in
    a single call. Never zero — one row at a time is the floor, however wide.
    """
    return max(1, min(SCAN_CHUNK_ROWS, SCAN_CHUNK_CELLS // max(max_col, 1)))
