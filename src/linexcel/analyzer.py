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

import formualizer as fz

from linexcel.decompose import (
    _collect_step_exprs,
    _decompose,
    _render_expr,  # noqa: F401  (re-exported: tests import it from analyzer)
)
from linexcel.engine import boot_engine
from linexcel.external import (
    find_workbooks,
    macro_files,
    read_external_links,
    resolve_books,
)
from linexcel.graph import (
    MAX_VALUE_SAMPLE,
    SMALL_RANGE_CELLS,
    GraphBuilder,
    _sample_range_values,
)
from linexcel.loader import _stepped, load_cached_values
from linexcel.powerquery import Query, QuerySource, read_queries
from linexcel.progress import Reporter
from linexcel.refs import (
    Rect,
    a1,
    parse_ref,
    parse_ref_detailed,
    stretch_ref,
)
from linexcel.resolver import (
    DEFAULT_STEP_SECONDS,
    MAX_SCRATCH_EVALS,
    _Budget,
    _collect_ref_strings,
    _external_warning,
    _ValueResolver,
)
from linexcel.structure import (
    inspect_workbook,  # noqa: F401  (re-exported: public API)
    read_structure,
)
from linexcel.sweep import FormulaGroup, sweep_sheets
from linexcel.tables import _build_table_index, _enrich_with_table
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
    # Reported one node at a time: this is where a dense workbook spends its
    # minutes, and where a run that looks stuck actually is. A phase that
    # prints only when it ends cannot tell anyone that.
    for node_id, grp in _stepped(
        reporter.phase("nodes+edges", total=len(kept_groups)),
        kept_groups,
        "building",
    ):
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
                    add_edge(ensure_opaque_node(ref), node_id, "dep")
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
            samples = [
                {
                    "addr": a1(r, c),
                    **resolver.describe(sheet, r, c, grp.formulas.get((r, c))),
                }
                for r, c in _spread_cells(grp.cells, MAX_VALUE_SAMPLE)
            ]

        steps = None
        # A volatile cell is shown as *not* recalculated, so decomposing it
        # would contradict its own card: every step under `=TODAY()+7` would
        # carry a figure computed from today's clock.
        if ast_dict is not None and value_fields.get("valueSource") != "volatile":
            # The root step is the formula itself: when the engine computed the
            # cell, its value is that step's value and needs no scratch pass.
            root_value = (
                value_fields["value"]
                if value_fields.get("valueSource") == "engine"
                else None
            )
            step_exprs = _collect_step_exprs(ast_dict, skip_root=root_value is not None)
            if step_exprs:
                resolver.preload_steps(step_exprs, sheet)
            steps = _decompose(
                ast_dict, sheet, resolver, defined_names, root_value=root_value
            )

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
        books = resolver.external_books(formula)
        if books:
            node["externalBooks"] = books
        _enrich_with_table(node, table_index, sheet, rep_r, rep_c)
        nodes[node_id] = node

    # --- 6. VBA --------------------------------------------------------------
    vba_modules = extract_vba_modules(data, filename, warnings)
    vba_procs: list[VbaProc] = analyze_vba(vba_modules) if vba_modules else []
    # Code a workbook calls often does not live in it: an .xlam add-in holds
    # the functions, and the workbook only names them. Given the folder, that
    # code is read too, and each module says which file it came from.
    if refs_dir is not None:
        for addin in macro_files(Path(refs_dir)):
            extra = extract_vba_modules(addin.read_bytes(), addin.name, warnings)
            if not extra:
                continue
            origin = {f"{addin.name}:{name}": code for name, code in extra.items()}
            vba_modules.update(origin)
            vba_procs.extend(analyze_vba(origin))
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


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


def _spread_cells(cells: list[tuple[int, int]], n: int) -> list[tuple[int, int]]:
    """``n`` cells spread evenly over a group, in reading order.

    The head of a stretch is made of near-identical neighbours: sampling
    ``B2, C2, B3`` out of 400,000 cells says nothing about the far end, which
    is exactly where a pattern that broke — a hard-coded cell dropped into the
    middle of a column — shows up. First and last are always included.
    """
    ordered = sorted(cells)
    if n < 2:
        return ordered[: max(n, 0)]
    if len(ordered) <= n:
        return ordered
    last = len(ordered) - 1
    picks = sorted({round(i * last / (n - 1)) for i in range(n)})
    return [ordered[i] for i in picks]


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
