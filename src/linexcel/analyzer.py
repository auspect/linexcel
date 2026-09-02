"""Builds the lineage graph of an Excel workbook.

Steps:
1. structure (sheets, dimensions, defined names) via openpyxl read-only;
2. formulas + computed values via the Rust engine formualizer;
3. grouping of stretched formulas by R1C1 canonicalization —
   a column of 50,000 copied formulas becomes ONE node;
4. resolution of precedents (cells, ranges, names, other sheets);
5. decomposition of each composite formula into individually evaluated steps
   in a scratch sheet of the engine;
6. lineage of extracted VBA code (oletools);
7. Power Query — the queries that fill a range no formula writes to.
"""

from __future__ import annotations

import datetime
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from linexcel.engine import boot_engine
from linexcel.external import find_workbooks, read_external_links, resolve_books
from linexcel.graph import GraphBuilder
from linexcel.loader import load_cached_values
from linexcel.powerquery import query_warning, read_queries
from linexcel.progress import Reporter
from linexcel.resolver import (
    DEFAULT_STEP_SECONDS,
    MAX_SCRATCH_EVALS,
    _Budget,
    _external_warning,
    _ValueResolver,
)
from linexcel.structure import (
    inspect_workbook,  # noqa: F401  (re-exported: public API)
    read_structure,
)
from linexcel.sweep import sweep_sheets
from linexcel.tables import _build_table_index


def analyze_workbook(
    data: bytes,
    filename: str = "workbook.xlsx",
    *,
    verbose: bool = False,
    refs_dir: str | Path | None = None,
    step_seconds: float | None = DEFAULT_STEP_SECONDS,
) -> dict[str, Any]:
    """Full analysis: returns the JSON-serializable graph and the engine.

    ``refs_dir`` is a folder holding the workbooks this one links to. Without
    it, a cell reading ``'[1]Annual'!B4`` is named and left unresolved — the
    engine has nothing to follow. With it, the referenced file is read and the
    reference is evaluated as the value it stands for; the report says, per
    workbook, whether it was read from that folder, from the cache Excel left
    in the file, or not at all.
    """
    warnings: list[str] = []
    _t0 = time.perf_counter()
    reporter = Reporter(verbose)

    def _v(label: str, t: float) -> None:
        reporter.note(f"{label}: {time.perf_counter() - t:.1f}s")

    # --- 1. structure -----------------------------------------------------
    _t = time.perf_counter()
    structure = read_structure(data)
    sheet_dims = structure.sheet_dims
    defined_names = structure.defined_names

    # Workbooks this one links to. Always named; read for real only when the
    # caller points at a folder holding them.
    externals = read_external_links(data)
    refs_files: dict[str, Path] = {}
    if refs_dir is not None:
        refs_files = find_workbooks(Path(refs_dir))
        if externals:
            resolve_books(externals, Path(refs_dir), warnings)
    _v("structure", _t)

    # values the file itself carries: last resort, and the only source of
    # dates and of what the user actually saw on screen
    _t = time.perf_counter()
    cached = load_cached_values(data, warnings, reporter)

    # --- 2. computation engine -------------------------------------------
    _t = time.perf_counter()
    session = boot_engine(data, sheet_dims, warnings)
    engine = session.engine
    engine_sheets = session.engine_sheets
    engine_alive = session.engine_alive
    quarantined = session.quarantined
    scratch_ready = session.scratch_ready
    _v("engine_init+evaluate_all", _t)

    # Tables: declared ones from the package parts, static ones from a small
    # window the engine already holds. A per-cell lookup enriching the nodes.
    _t = time.perf_counter()
    table_index = _build_table_index(data, engine, sheet_dims, engine_sheets)
    _v("tables", _t)

    budget = _Budget(MAX_SCRATCH_EVALS, step_seconds)
    resolver = _ValueResolver(
        engine,
        engine_sheets,
        cached,
        warnings,
        budget,
        scratch_ready,
        engine_alive=engine_alive,
        sheet_dims=sheet_dims,
        externals=externals,
        refs_files=refs_files,
    )

    # --- 3. extraction + grouping ------------------------------------------
    sweep = sweep_sheets(
        engine, sheet_dims, engine_sheets, quarantined, warnings, reporter
    )
    groups = sweep.groups
    formula_count = sweep.formula_count
    sheet_stats = sweep.sheet_stats

    # --- 4. nodes + edges: names, formulas, VBA, Power Query ---------------
    _t = time.perf_counter()
    builder = GraphBuilder(
        resolver, sheet_dims, table_index, defined_names, warnings, reporter
    )
    builder.select_nodes(groups)
    nodes = builder.nodes
    edges = builder.edges
    kept_groups = builder.kept_groups
    builder.build_names()
    builder.build_formula_nodes()

    # --- 5. VBA (oletools) ---------------------------------------------------
    builder.build_vba(data, filename, refs_dir)
    vba_modules = builder.vba_modules
    vba_procs = builder.vba_procs

    # --- 6. Power Query -------------------------------------------------------
    # A range filled by a query has no formula above it, so without this the
    # graph shows where the data landed and nothing about where it came from.
    queries = read_queries(data)
    builder.build_queries(queries)

    pq_warning = query_warning(queries)
    if pq_warning:
        warnings.append(pq_warning)

    if not engine_alive and resolver.n_recovered + resolver.n_unrecovered:
        warnings.append(
            f"Values recovered cell by cell: {resolver.n_recovered} recomputed, "
            f"{resolver.n_unrecovered} left to the value stored in the file"
        )
    uncomputed = resolver.uncomputed_warning()
    if uncomputed:
        warnings.append(uncomputed)
    external_warning = _external_warning(resolver.external_workbooks(), refs_dir)
    if external_warning:
        warnings.append(external_warning)

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
                "externalWorkbooks": len(
                    {b.name for b in resolver.external_workbooks() if b.name}
                ),
                "externalWorkbooksRead": len(
                    {b.name for b in resolver.external_workbooks() if b.resolved}
                ),
                "queries": len(queries),
                "queriesLoaded": sum(1 for q in queries if q.loaded),
            },
        },
        "sheets": list(sheet_dims.keys()),
        "nodes": list(nodes.values()),
        "edges": list(edges.values()),
    }
    exhausted = budget.warning()
    if exhausted:
        warnings.append(exhausted)
    _v("graph", _t)
    if verbose:
        print(
            f"[linexcel] total: {time.perf_counter() - _t0:.1f}s | "
            f"{len(nodes)} nodes | {len(edges)} edges | "
            f"{formula_count:,} formulas",
            file=sys.stderr,
        )
    return {"graph": graph, "engine": engine, "analysisId": uuid.uuid4().hex[:16]}
