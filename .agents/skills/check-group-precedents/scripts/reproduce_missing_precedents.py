#!/usr/bin/env python3
"""Show the known finding on a real workbook, through `linexcel.analyze`.

The exhaustive check works on the functions. This script shows what the same
defect does to a user's report: it writes a ten-row workbook, analyzes it and
compares the cells the formulas read with the cells the group node is wired to.

    uv run python <this file>

Exit 3 when cells are missing from the lineage (the finding is present),
0 when every cell read has an edge (the finding is fixed).
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook

from linexcel import analyze
from linexcel.refs import parse_ref

ROOT = Path(__file__).resolve().parents[4]
OUT = ROOT / "validation_screenshots" / "check-group-precedents"
ROWS = 10
ANCHOR = 5


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        getattr(stream, "reconfigure", lambda **_: None)(encoding="utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "anchored_end_crossed.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    for row in range(1, ROWS + 1):
        sheet.cell(row=row, column=1, value=row)
        # What a generator writes when it shifts the relative end row by row.
        # Excel itself would store =SUM(A$5:$A6) from row 6 on.
        sheet.cell(row=row, column=2, value=f"=SUM(A{row}:$A${ANCHOR})")
    book.save(path)

    graph = analyze(path).graph
    groups = [n for n in graph["nodes"] if n.get("kind") == "group"]
    print(f"workbook: {path}")
    print(
        f"B1:B{ROWS} hold =SUM(A1:$A${ANCHOR}) to =SUM(A{ROWS}:$A${ANCHOR}), "
        "the relative end shifted row by row"
    )
    read = {(row, 1) for row in range(1, ROWS + 1)}
    wired: set[tuple[int, int]] = set()
    for node in groups:
        print(
            f"group node {node['id']}  cells {node.get('bbox')}  "
            f"formula {node.get('formula')}"
        )
        for edge in graph["edges"]:
            if edge["target"] != node["id"]:
                continue
            rect = parse_ref(edge["source"].split(":", 1)[1])
            print(f"  precedent {edge['source']}")
            if rect is not None:
                wired |= {
                    (row, col)
                    for row in range(rect.r1, min(rect.r2, ROWS) + 1)
                    for col in range(rect.c1, min(rect.c2, 2) + 1)
                }
    missing = sorted(read - wired)
    if missing:
        names = ", ".join(f"A{row}" for row, _col in missing)
        print(f"MISSING from the lineage: {names}")
        print(
            "These cells are read by the formulas below the anchor row and "
            "have no edge to the group."
        )
        return 3
    print("every cell read by the group has an edge to it")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
