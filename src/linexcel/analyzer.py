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
from collections import defaultdict
from pathlib import Path
from typing import Any

from linexcel.decompose import (
    _render_expr,  # noqa: F401  (re-exported: tests import it from analyzer)
)
from linexcel.engine import boot_engine
from linexcel.external import (
    find_workbooks,
    macro_files,
    read_external_links,
    resolve_books,
)
from linexcel.graph import SMALL_RANGE_CELLS, GraphBuilder
from linexcel.loader import load_cached_values
from linexcel.powerquery import Query, QuerySource, read_queries
from linexcel.progress import Reporter
from linexcel.refs import Rect, parse_ref, parse_ref_detailed
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
from linexcel.vba import VbaProc, analyze_vba, extract_vba_modules

MAX_VBA_CODE_CHARS = 6_000
MAX_QUERY_CODE_CHARS = 6_000
#: How many query sources one warning line names before it says "and more".
MAX_QUERY_SOURCES_SHOWN = 6


def _query_warning(queries: list[Query]) -> str | None:
    """One line for the queries that feed the workbook from outside it.

    A query whose source is a file, a URL or a server is a dependency of the
    same nature as a link to another workbook: the values it produced are in
    the file, what produced them is not. The graph shows the query and names
    its source; nobody should read that as the source having been checked.
    """
    if not queries:
        return None
    loaded = sum(1 for query in queries if query.loaded)
    if len(queries) == 1:
        head = "1 Power Query query feeds this workbook" + (
            ", loaded onto a sheet."
            if loaded
            else ", loaded nowhere (connection only)."
        )
    else:
        head = (
            f"{len(queries)} Power Query queries feed this workbook, "
            f"{loaded} of them loaded onto a sheet."
        )
    parts = [head]
    outside = sorted(
        {source.target for query in queries for source in query.outside_sources()}
    )
    if outside:
        shown = ", ".join(outside[:MAX_QUERY_SOURCES_SHOWN])
        if len(outside) > MAX_QUERY_SOURCES_SHOWN:
            shown += f", … (+{len(outside) - MAX_QUERY_SOURCES_SHOWN})"
        parts.append(
            f"Their data comes from outside the file and was not read: {shown}."
        )
    return " ".join(parts)


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

    # --- 3. extraction + grouping ----------------------------------------
    sweep = sweep_sheets(engine, sheet_dims, engine_sheets, quarantined, warnings, reporter)
    groups = sweep.groups
    formula_count = sweep.formula_count
    sheet_stats = sweep.sheet_stats

    # --- 4. formula nodes -------------------------------------------------
    _t = time.perf_counter()
    builder = GraphBuilder(resolver, sheet_dims, table_index, defined_names, warnings, reporter)
    builder.select_nodes(groups)
    nodes = builder.nodes
    edges = builder.edges
    input_nodes = builder.input_nodes
    cell_owner = builder.cell_owner
    ast_cache = builder.ast_cache
    kept_groups = builder.kept_groups
    ensure_opaque_node = builder.ensure_opaque_node
    ensure_input_node = builder.ensure_input_node
    add_edge = builder.add_edge
    resolve_rect_edges = builder.resolve_rect_edges

    # defined names -----------------------------------------------------------
    builder.build_names()
    name_nodes = builder.name_nodes

    # formula nodes + edges -------------------------------------------------
    builder.build_formula_nodes()

    # --- 6. VBA --------------------------------------------------------------
    builder.build_vba(data, filename, refs_dir)
    vba_modules = builder.vba_modules
    vba_procs = builder.vba_procs

    # --- 7. Power Query ------------------------------------------------------
    # A range filled by a query has no formula above it, so without this the
    # graph shows where the data landed and nothing about where it came from.
    queries = read_queries(data)
    # Keyed on the exact name: M is case-sensitive, so folding here would let
    # two queries that differ only in case collapse onto one node.
    query_ids = {q.name: f"q:{q.name}" for q in queries}
    tables_by_name = {
        table["name"].casefold(): (sheet, table["ref"])
        for sheet, entries in table_index.items()
        for table in entries
        if table.get("name") and table.get("ref")
    }

    def query_source_node(source: QuerySource) -> str:
        """A node for something a query reads that is not in this workbook."""
        node_id = ensure_input_node(Rect(None, 1, 1, 1, 1), opaque_label=source.target)
        nodes[node_id].setdefault("sourceKind", source.kind)
        nodes[node_id].setdefault("function", source.function)
        return node_id

    for query in queries:
        qid = query_ids[query.name]
        nodes[qid] = {
            "id": qid,
            "kind": "query",
            "label": query.name,
            "sheet": query.loaded_to[0].sheet if query.loaded_to else None,
            "code": query.source[:MAX_QUERY_CODE_CHARS],
            "loadedTo": [d.as_dict() for d in query.loaded_to],
            "sources": [s.as_dict() for s in query.sources],
        }
    for query in queries:
        qid = query_ids[query.name]
        for source in query.sources:
            if source.kind == "query":
                upstream = query_ids.get(source.target)
                if upstream is not None:
                    add_edge(upstream, qid, "query")
                continue
            if source.kind == "table":
                # ``Excel.CurrentWorkbook`` reads a table or a defined name of
                # this very file: that end of the link is in the graph already.
                placed = tables_by_name.get(source.target.casefold())
                if placed is not None:
                    rect = parse_ref(placed[1], default_sheet=placed[0])
                    if rect is not None:
                        resolve_rect_edges(rect, qid, kind="query")
                        continue
                named = name_nodes.get(source.target.upper())
                if named is not None:
                    add_edge(named, qid, "query")
                    continue
            add_edge(query_source_node(source), qid, "query")
        for destination in query.loaded_to:
            rect = (
                parse_ref(destination.ref, default_sheet=destination.sheet)
                if destination.ref
                else None
            )
            if rect is not None:
                add_edge(qid, ensure_input_node(rect), "query-load")

    query_warning = _query_warning(queries)
    if query_warning:
        warnings.append(query_warning)

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
