"""Orchestrates the phases that build the lineage graph of an Excel workbook:
structure → external links → cached values → engine boot → tables →
resolver → formula sweep → graph nodes/edges (names, formulas, VBA, Power
Query) → assembly. Each phase is a leaf module; this file only sequences
them and assembles the result.
"""

from __future__ import annotations

import datetime
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from linexcel.engine import CHAIN_LAYERS_WARNING, boot_engine
from linexcel.external import find_workbooks, read_external_links, resolve_books
from linexcel.graph import GraphBuilder
from linexcel.loader import load_cached_values
from linexcel.powerquery import query_warning, read_queries
from linexcel.progress import Reporter
from linexcel.refs import Rect, a1, parse_ref
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


def _parse_targets(targets: list[str]) -> list[tuple[str, int, int]]:
    """Sheet-qualified single cells, in the order given, duplicates removed.

    Entries may carry several comma-separated cells, so the CLI can pass its
    repeated ``--target`` flags straight through. A target without a sheet is
    ambiguous in a multi-sheet workbook, and a range is not a target but a
    set of them — spell those as several targets rather than one span.
    """
    cells: list[tuple[str, int, int]] = []
    for chunk in targets:
        for raw in (t.strip() for t in chunk.split(",")):
            if not raw:
                continue
            rect = parse_ref(raw)
            if (rect is None or rect.sheet is None) and "!" in raw:
                # A sheet name with a space is legal unquoted on the command
                # line — 'My Sheet'!B1 is the formula spelling, not the
                # user's. Split on the last '!': the sheet's existence is
                # checked against the workbook later either way.
                sheet, _, body = raw.rpartition("!")
                rect = parse_ref(body, default_sheet=sheet.strip("'"))
            if rect is None or rect.sheet is None or rect.ncells != 1:
                raise ValueError(
                    f"Invalid target {raw!r}: expected a sheet-qualified cell, "
                    f"like 'Sheet1!A1'."
                )
            cell = (rect.sheet, rect.r1, rect.c1)
            if cell not in cells:
                cells.append(cell)
    return cells


def _longest_dep_chain(nodes: dict[str, Any], edges: dict) -> int:
    """Longest precedent→dependent path across formula nodes, in steps.

    Iterative post-order walk: a chain of running totals can be tens of
    thousands of nodes deep, past any safe recursion limit. A cycle
    contributes nothing — the engine stamps its members ``#CIRC!`` and the
    graph reports them elsewhere.
    """
    cellish = {
        nid
        for nid, node in nodes.items()
        if node.get("kind") in {"cell", "group", "input", "misc"}
    }
    adjacency: dict[str, set[str]] = {}
    for edge in edges.values():
        src, dst = edge["source"], edge["target"]
        if edge["kind"] == "dep" and src in cellish and dst in cellish:
            adjacency.setdefault(src, set()).add(dst)
    depth: dict[str, int] = {}
    best = 0
    for root in adjacency:
        if root in depth:
            continue
        on_stack = {root}
        stack = [(root, False)]
        while stack:
            node, expanded = stack.pop()
            if expanded:
                d = 1 + max(
                    (depth.get(child, 0) for child in adjacency.get(node, ())),
                    default=0,
                )
                depth[node] = d
                best = max(best, d)
                on_stack.discard(node)
            elif node not in depth and node not in on_stack:
                on_stack.add(node)
                stack.append((node, True))
                stack.extend(
                    (child, False)
                    for child in adjacency.get(node, ())
                    if child not in depth and child not in on_stack
                )
    return best


def _chain_depth_warning(
    nodes: dict[str, Any], edges: dict, intra_chain: int
) -> str | None:
    """Flag a workbook whose formulas depend on each other in long chains.

    The pre-run estimate weighs formulas and counts them but cannot see how
    deep the dependency chains run, and the engine's own evaluation plan only
    counts layers while cells are still dirty — after ``evaluate_all`` it
    reads zero. So the measure is taken from the lineage itself: the longest
    path across formula nodes, and the largest self-referencing group (a
    running-total column is one node on the graph but a chain as deep as it
    is long). An indicator of risk, not a duration estimate.
    """
    chain = max(_longest_dep_chain(nodes, edges), intra_chain)
    if chain < CHAIN_LAYERS_WARNING:
        return None
    return (
        f"Long dependency chains: up to {chain:,} linked calculation steps. "
        f"Such workbooks are where the step-by-step decomposition hits its "
        f"time budget (raise it with --time-budget) and where an analysis "
        f"runs long — a risk indicator, not a duration estimate"
    )


def analyze_workbook(
    data: bytes,
    filename: str = "workbook.xlsx",
    *,
    verbose: bool = False,
    refs_dir: str | Path | None = None,
    step_seconds: float | None = DEFAULT_STEP_SECONDS,
    targets: list[str] | None = None,
) -> dict[str, Any]:
    """Full analysis: returns the JSON-serializable graph and the engine.

    ``refs_dir`` is a folder holding the workbooks this one links to. Without
    it, a cell reading ``'[1]Annual'!B4`` is left unresolved; with it, the
    reference is read and evaluated.

    ``targets`` limits the analysis to the upstream subgraph of those cells
    (``["Sheet1!A1", ...]``): the engine boots without a global evaluation,
    only the cells feeding the targets are traced, evaluated and graphed, and
    the rest of the workbook is omitted from the lineage rather than
    evaluated. Without it the whole workbook is analysed, as before.
    """
    warnings: list[str] = []
    _t0 = time.perf_counter()
    reporter = Reporter(verbose)
    target_cells = _parse_targets(targets) if targets else None

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
    session = boot_engine(data, warnings, reporter, targets=target_cells)
    engine = session.engine
    engine_sheets = session.engine_sheets
    engine_alive = session.engine_alive
    quarantined = session.quarantined
    scratch_ready = session.scratch_ready
    reachable = session.reachable

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
        engine,
        sheet_dims,
        engine_sheets,
        quarantined,
        warnings,
        reporter,
        reachable=reachable,
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
    if target_cells:
        # A target holding a constant has no formula group to own it; it
        # still gets a node, so the subgraph the user asked for has its root.
        for sheet, row, col in target_cells:
            if (row, col) not in builder.cell_owner.get(sheet, {}):
                builder.ensure_input_node(Rect(sheet, row, col, row, col))
        asked = ", ".join(f"{s}!{a1(r, c)}" for s, r, c in target_cells)
        warnings.append(
            f"Targeted analysis of {asked}: only the {len(reachable or []):,} "
            f"cell(s) of the upstream subgraph were evaluated and graphed; "
            f"the rest of the workbook is omitted from the lineage and its "
            f"formulas were not recomputed"
        )
    if engine_alive and not target_cells:
        # In targeted mode the engine's own evaluation plan already flagged
        # the subgraph, exactly, before evaluation.
        chain_warning = _chain_depth_warning(nodes, edges, builder.intra_chain)
        if chain_warning:
            warnings.append(chain_warning)

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
        **(
            {"targets": [f"{s}!{a1(r, c)}" for s, r, c in target_cells]}
            if target_cells
            else {}
        ),
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
