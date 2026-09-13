"""Sweep formula cells sheet by sheet and group them by R1C1 pattern.

Extracted mechanically from analyzer.py: read formulas in bounded chunks,
re-inject the formulas quarantined by ``engine.boot_engine`` so they still
show in the lineage, and group cells sharing the same R1C1-canonicalized
formula into one FormulaGroup — a column of 50,000 copied formulas becomes
ONE entry.

A targeted analysis (``reachable`` set by ``boot_engine``) sweeps only the
rows the upstream subgraph touches and keeps only its cells; the rest of the
workbook is omitted from the lineage rather than swept and dropped.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

import formualizer as fz

from linexcel.engine import _chunk_rows
from linexcel.loader import MAX_CELLS_PER_SHEET
from linexcel.progress import Reporter
from linexcel.rewrite import canonical_r1c1


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


@dataclass
class SweepResult:
    groups: dict[tuple[str, str], FormulaGroup]
    formula_count: int
    sheet_stats: list[dict[str, Any]]


def _target_ranges(cells: list[tuple[int, int]]):
    """Rectangles containing only requested cells, with bounded dense reads.

    Merge adjacent columns and then identical spans on adjacent rows. A dense
    target column stays one batch; distant cells never pay for the gaps.
    """
    by_row: dict[int, list[int]] = defaultdict(list)
    for row, col in cells:
        by_row[row].append(col)
    pending: dict[tuple[int, int], tuple[int, int]] = {}
    for row in sorted(by_row):
        columns = sorted(by_row[row])
        spans = []
        first = last = columns[0]
        for col in columns[1:]:
            if col == last + 1:
                last = col
            else:
                spans.append((first, last))
                first = last = col
        spans.append((first, last))
        current = {}
        for c1, c2 in spans:
            previous = pending.pop((c1, c2), None)
            if previous is not None:
                r1, r2 = previous
                if r2 == row - 1 and row - r1 < _chunk_rows(c2 - c1 + 1):
                    current[(c1, c2)] = (r1, row)
                    continue
                yield r1, c1, r2, c2
            current[(c1, c2)] = (row, row)
        for (c1, c2), (r1, r2) in pending.items():
            yield r1, c1, r2, c2
        pending = current
    for (c1, c2), (r1, r2) in pending.items():
        yield r1, c1, r2, c2


def sweep_sheets(
    engine,
    sheet_dims: dict[str, tuple[int, int]],
    engine_sheets: set[str],
    quarantined: dict[tuple[str, int, int], str],
    warnings: list[str],
    reporter: Reporter,
    *,
    reachable: set[tuple[str, int, int]] | None = None,
    max_cells_per_sheet: int | None = None,
) -> SweepResult:
    cells_limit = (
        MAX_CELLS_PER_SHEET
        if max_cells_per_sheet is None
        else max_cells_per_sheet
    )
    groups: dict[tuple[str, str], FormulaGroup] = {}
    formula_count = 0
    sheet_stats: list[dict[str, Any]] = []
    wanted_by_sheet: dict[str, list[tuple[int, int]]] = defaultdict(list)
    if reachable is not None:
        for sheet, row, col in reachable:
            wanted_by_sheet[sheet].append((row, col))

    def record(sheet: str, r: int, c: int, formula: str | None) -> int:
        formula = formula or quarantined.get((sheet, r, c))
        if not formula:
            return 0
        key = (sheet, canonical_r1c1(formula, r, c))
        grp = groups.get(key)
        if grp is None:
            grp = groups[key] = FormulaGroup(sheet, key[1])
        grp.cells.append((r, c))
        if len(grp.formulas) < 3:
            grp.formulas[(r, c)] = formula
        elif (r, c) < max(grp.formulas):
            del grp.formulas[max(grp.formulas)]
            grp.formulas[(r, c)] = formula
        return 1

    with reporter.phase("extraction+grouping", total=len(sheet_dims)) as _bar:
        for sheet, (max_row, max_col) in sheet_dims.items():
            if sheet not in engine_sheets:
                warnings.append(f"Sheet '{sheet}' skipped (not loaded by engine)")
                continue
            if reachable is not None:
                n_formulas = 0
                fsheet = engine.sheet(sheet)
                for r1, c1, r2, c2 in _target_ranges(wanted_by_sheet[sheet]):
                    try:
                        rows = fsheet.get_formulas(
                            fz.RangeAddress(sheet, r1, c1, r2, c2)
                        )
                    except Exception as exc:
                        warnings.append(f"Could not read formulas on {sheet}: {exc}")
                        continue
                    for i, row_vals in enumerate(rows):
                        for j, formula in enumerate(row_vals):
                            n_formulas += record(sheet, r1 + i, c1 + j, formula)
                formula_count += n_formulas
                sheet_stats.append(
                    {
                        "name": sheet,
                        "rows": max_row,
                        "cols": max_col,
                        "formulaCells": n_formulas,
                    }
                )
                _bar.step(f"sweeping {sheet}")
                continue
            n_formulas = 0
            scanned = 0
            fsheet = engine.sheet(sheet)
            chunk_rows = _chunk_rows(max_col)
            r0 = 1
            while r0 <= max_row:
                # The ceiling is spent in rows, and the last chunk is clipped to
                # what is left rather than dropped whole: dropping it stopped a
                # 4,000,000-cell budget at 3,600,000 and lost every row of the
                # chunk that would have overshot.
                rows_left = (cells_limit - scanned) // max_col
                if rows_left <= 0:
                    warnings.append(
                        f"Sheet '{sheet}' scanned to row {r0 - 1:,} of {max_row:,} "
                        f"({cells_limit:,} cell ceiling): formulas below "
                        f"that row are missing from the lineage"
                    )
                    break
                r1 = min(r0 + chunk_rows - 1, max_row, r0 + rows_left - 1)
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
                        n_formulas += record(sheet, r, j + 1, f)
                r0 = r1 + 1
            formula_count += n_formulas
            sheet_stats.append(
                {
                    "name": sheet,
                    "rows": max_row,
                    "cols": max_col,
                    "formulaCells": n_formulas,
                }
            )
            _bar.step(f"sweeping {sheet}")

    return SweepResult(groups, formula_count, sheet_stats)
