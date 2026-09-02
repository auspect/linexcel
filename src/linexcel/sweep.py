"""Sweep formula cells sheet by sheet and group them by R1C1 pattern.

Extracted mechanically from analyzer.py: read formulas in bounded chunks,
re-inject the formulas quarantined by ``engine.boot_engine`` so they still
show in the lineage, and group cells sharing the same R1C1-canonicalized
formula into one FormulaGroup — a column of 50,000 copied formulas becomes
ONE entry.
"""

from __future__ import annotations

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


def sweep_sheets(
    engine,
    sheet_dims: dict[str, tuple[int, int]],
    engine_sheets: set[str],
    quarantined: dict[tuple[str, int, int], str],
    warnings: list[str],
    reporter: Reporter,
) -> SweepResult:
    groups: dict[tuple[str, str], FormulaGroup] = {}
    formula_count = 0
    sheet_stats: list[dict[str, Any]] = []

    with reporter.phase("extraction+grouping", total=len(sheet_dims)) as _bar:
        for sheet, (max_row, max_col) in sheet_dims.items():
            if sheet not in engine_sheets:
                warnings.append(f"Sheet '{sheet}' skipped (not loaded by engine)")
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
                rows_left = (MAX_CELLS_PER_SHEET - scanned) // max_col
                if rows_left <= 0:
                    warnings.append(
                        f"Sheet '{sheet}' scanned to row {r0 - 1:,} of {max_row:,} "
                        f"({MAX_CELLS_PER_SHEET:,} cell ceiling): formulas below "
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
