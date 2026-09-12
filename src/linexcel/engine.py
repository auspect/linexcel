"""Boot the formualizer engine, and make it survive a broken reference.

Extracted mechanically from analyzer.py: instantiate the engine from the
workbook bytes, run the whole-workbook evaluation, and — when that fails
because one formula names something the engine cannot resolve — quarantine
the offending cells and retry, so the rest of the file keeps its computed
values.
"""

from __future__ import annotations

import io
import math
import re
import threading
import zipfile
from bisect import bisect_left, bisect_right
from collections import deque
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree
from xml.sax.saxutils import unescape

import formualizer as fz

from linexcel.decompose import SCRATCH_SHEET
from linexcel.loader import _detect_epoch_1904
from linexcel.progress import Reporter
from linexcel.refs import a1, col_to_num, parse_ref, quote_sheet

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

#: A cell element and its body. ``<c>`` never nests another ``<c>``, so the
#: lazy ``.*?</c>`` is safe; self-closing ``<c/>`` elements carry no formula.
_CELL_RE = re.compile(rb'<c\b[^>]*\br="([A-Z]{1,3})(\d+)"[^>]*(?<!/)>(.*?)</c>', re.S)
#: The ``<f>`` inside a cell body: either ``<f …/>`` (a shared-formula slave)
#: or ``<f …>text</f>``.
_F_RE = re.compile(rb"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
#: The shared-formula index, ``si="3"``.
_SI_RE = re.compile(rb'\bsi="(\d+)"')
_XML_ENTITIES = {"&quot;": '"', "&apos;": "'"}

#: Deepest formula parse tree ``evaluate_all`` is trusted with. The evaluator
#: walks the tree recursively, and past roughly 1,000 nested operations that
#: walk overflows the stack and *aborts the process* — a hard exit no
#: try/except sees (upstream: https://github.com/PSU3D0/formualizer/issues/411;
#: measured here on 0.9.3: a 700-term chain evaluates, a 1,000-term one
#: aborts). Formulas deeper than this are quarantined before the evaluation
#: runs. Configurable per call via ``boot_engine``'s ``max_ast_depth``.
MAX_AST_DEPTH = 900

#: Whether the engine evaluates independent branches of the dependency graph
#: in parallel. This is formualizer's own default since 0.9; it is stated
#: explicitly so the choice is documented and so a workbook that misbehaves
#: under parallel evaluation has an off switch — ``boot_engine(parallel=False)``
#: — instead of a patched engine. Prototype timings live in
#: ``scripts/perf_probe.py``.
PARALLEL_EVALUATION = True

#: Budgets for the upstream trace behind targeted evaluation. The defaults
#: (depth 6, 512 nodes) are an interactive-explorer budget, not an analysis
#: one: a running-total column is a chain exactly as deep as it is long.
#: A trace that still hits one of these reports it through its truncation
#: flag, and the warning names the analysis as partial.
TRACE_MAX_DEPTH = 100_000
TRACE_MAX_NODES = 2_000_000
TRACE_MAX_LINKS = 4_000_000
TRACE_MAX_WORK = 20_000_000
TRACE_RANGE_MEMBERS = 100_000

#: From how many linked calculation steps a workbook counts as chain-heavy.
#: An indicator of risk, not a duration estimate: what long chains do to the
#: step-by-step decomposition is the warning's business, not the stopwatch's.
CHAIN_LAYERS_WARNING = 25


@dataclass
class EngineSession:
    engine: Any
    engine_sheets: set[str]
    engine_alive: bool
    quarantined: dict[tuple[str, int, int], str]
    scratch_ready: bool
    #: Cells of the upstream subgraph, set only by a targeted boot; ``None``
    #: means the whole workbook was evaluated, as before.
    reachable: set[tuple[str, int, int]] | None = None
    unavailable: set[tuple[str, int, int]] = field(default_factory=set)


def _open_workbook(data: bytes, parallel: bool):
    """Instantiate the engine with the configured evaluation options."""
    eval_config = fz.EvaluationConfig()
    eval_config.enable_parallel = parallel
    # The importer uses this option while converting formatted numeric cells
    # into dates; setting it after loading cannot undo a 1,462-day shift.
    eval_config.date_system = "1904" if _detect_epoch_1904(data) else "1900"
    iterate, count, delta = _iteration_options(data)
    eval_config.cycle_policy = "iterate" if iterate else "error"
    eval_config.iterate_max_iterations = count
    eval_config.iterate_max_change = delta
    workbook = fz.Workbook.from_bytes(
        data, config=fz.WorkbookConfig(eval_config=eval_config)
    )
    _register_legacy_normal_functions(workbook, eval_config.date_system)
    return workbook


def _register_legacy_normal_functions(workbook, date_system: str) -> None:
    """Bridge two Excel compatibility names to their native modern functions.

    Registering workbook-local functions preserves original formulas, traces
    and scratch expressions. The callback must never re-enter its workbook;
    formualizer explicitly permits using a separate workbook. That evaluator
    owns only argument cells and delegates coercion and errors to the native
    statistical functions, instead of approximating their Excel semantics.
    """
    native = None
    lock = threading.Lock()

    def evaluate(modern: str, arguments: tuple) -> Any:
        nonlocal native
        with lock:
            if native is None:
                config = fz.EvaluationConfig()
                config.enable_parallel = False
                config.date_system = date_system
                native = fz.Workbook(config=fz.WorkbookConfig(eval_config=config))
                for name in ("Result", "Arg1", "Arg2", "Arg3", "Arg4"):
                    native.add_sheet(name)
            references = []
            for index, argument in enumerate(arguments, 1):
                sheet = f"Arg{index}"
                if isinstance(argument, list):
                    rows = argument or [[None]]
                    height, width = len(rows), len(rows[0])
                    native.sheet(sheet).set_values_batch(1, 1, height, width, rows)
                    references.append(f"{sheet}!A1:{a1(height, width)}")
                else:
                    native.set_value(sheet, 1, 1, argument)
                    references.append(f"{sheet}!A1")
            if modern == "NORM.S.DIST":
                references.append("TRUE")
            native.set_formula("Result", 1, 1, f"={modern}({','.join(references)})")
            return native.evaluate_cell("Result", 1, 1)

    workbook.register_function(
        "NORMSDIST",
        lambda *args: evaluate("NORM.S.DIST", args),
        min_args=1,
        max_args=1,
        thread_safe=False,
    )
    workbook.register_function(
        "NORMDIST",
        lambda *args: evaluate("NORM.DIST", args),
        min_args=4,
        max_args=4,
        thread_safe=False,
    )


def _iteration_options(data: bytes) -> tuple[bool, int, float]:
    """Workbook-declared iteration settings, with Excel's default limits."""
    enabled, count, delta = False, 100, 0.001
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            properties = ElementTree.fromstring(archive.read("xl/workbook.xml")).find(
                "{*}calcPr"
            )
        if properties is not None:
            enabled = properties.get("iterate", "").strip() in {"1", "true"}
            count = int(properties.get("iterateCount", "100"))
            delta = float(properties.get("iterateDelta", "0.001"))
    except (KeyError, ValueError, ElementTree.ParseError, zipfile.BadZipFile):
        return enabled, 100, 0.001
    return (
        enabled,
        count if count > 0 else 100,
        delta if math.isfinite(delta) and delta >= 0 else 0.001,
    )


def boot_engine(
    data: bytes,
    warnings: list[str],
    reporter: Reporter | None = None,
    *,
    max_ast_depth: int = MAX_AST_DEPTH,
    parallel: bool = PARALLEL_EVALUATION,
    targets: list[tuple[str, int, int]] | None = None,
) -> EngineSession:
    """Instantiate the engine and run its evaluation.

    ``evaluate_all`` is all-or-nothing, and it gives up on the *first*
    reference it cannot resolve — so a single formula pointing at another
    workbook costs every other cell in the file its computed value. When that
    happens, the offending formulas are cut out of the package bytes and the
    pass retried on the sanitized copy, leaving only them to the slower
    per-cell recovery.

    Two formulas never reach ``evaluate_all`` at all: the ones whose parse
    tree is deeper than ``max_ast_depth``, because the evaluator's recursive
    walk aborts the process on them rather than raising. They are quarantined
    up front, before the first evaluation.

    With ``targets`` the global pass is skipped entirely: the upstream
    subgraph of those cells is traced, only it is evaluated, and the session
    carries the reachable set so the sweep can leave the rest of the
    workbook out of the lineage.

    This is the long silent stretch of a large workbook — it used to be the
    one phase nothing reported while it ran — so it is a reporter phase like
    the others, with a step per stage rather than a timing printed after.
    """
    reporter = reporter or Reporter()
    with reporter.phase("engine evaluation") as progress:
        progress.step("loading workbook")
        data, chartsheets = _without_chartsheets(data)
        if chartsheets:
            warnings.append(
                f"Chartsheet(s) skipped: {', '.join(chartsheets)}. A chartsheet "
                f"holds a chart and no cells, and the engine refuses a whole "
                f"workbook that contains one"
            )
        engine = _open_workbook(data, parallel)
        engine_sheets = set(engine.sheet_names)
        engine_alive = True
        quarantined: dict[tuple[str, int, int], str] = {}
        too_deep, deepest = _find_too_deep(data, engine_sheets, max_ast_depth)
        unavailable: set[tuple[str, int, int]] = set()
        if too_deep:
            progress.step(f"quarantining {len(too_deep)} over-deep formula(s)")
            quarantined = too_deep
            engine = _open_workbook(
                _blank_formulas_in_package(data, quarantined), parallel
            )
            unavailable = _uncached_dependents(engine, data, too_deep, warnings)
            _mark_uncached_quarantine(engine, quarantined)
            sheet, row, col = deepest
            warnings.append(
                f"{len(too_deep)} cell(s) hold a formula nested deeper than the "
                f"engine can safely evaluate (parse tree over {max_ast_depth} "
                f"levels; evaluating one has aborted the process outright on such "
                f"input — https://github.com/PSU3D0/formualizer/issues/411). They "
                f"were quarantined before evaluation and keep the value stored in "
                f"the file, if any; without a cached value, dependent formulas "
                f"also remain uncomputed. Deepest: "
                f"{sheet}!{a1(row, col)}"
            )
        if targets is not None:
            reachable, target_alive = _evaluate_targets(
                engine, engine_sheets, targets, warnings, progress
            )
            unavailable |= _cycle_unavailable(
                engine,
                data,
                engine_sheets,
                warnings,
                reachable,
                evaluation_failed=not target_alive,
            )
            scratch_ready = _ensure_scratch(engine)
            return EngineSession(
                engine,
                engine_sheets,
                target_alive,
                quarantined,
                scratch_ready,
                reachable,
                unavailable,
            )
        try:
            progress.step("evaluating formulas")
            engine.evaluate_all()
        except Exception as exc:  # graph remains useful without values
            unresolvable = _find_unresolvable(data, engine_sheets)
            retried = False
            if unresolvable:
                quarantined |= unresolvable
                progress.step(f"retrying without {len(quarantined)} blocked cell(s)")
                try:
                    engine = _open_workbook(
                        _blank_formulas_in_package(data, quarantined), parallel
                    )
                    _mark_uncached_quarantine(engine, quarantined)
                    engine.evaluate_all()
                    retried = True
                except Exception:
                    pass
            if retried:
                warnings.append(
                    f"Global evaluation completed after isolating {len(quarantined)} "
                    f"cell(s) whose references the engine cannot resolve. Their "
                    f"stored values are used where available; dependent formulas "
                    f"without sufficient input remain uncomputed. First blocker: {exc}"
                )
            else:
                warnings.append(f"Global evaluation incomplete: {exc}")
                # Values are recovered cell by cell further down.
                engine_alive = False
                if _formulas_gone(engine, data, engine_sheets, quarantined):
                    # formualizer 0.9.3 keeps the formula map readable after a
                    # failed evaluate_all, so the rebuild is usually wasted —
                    # one full from_bytes. It is paid only when the probe says
                    # the engine really did come back empty.
                    rebuilt_data = (
                        _blank_formulas_in_package(data, too_deep) if too_deep else data
                    )
                    engine = _open_workbook(rebuilt_data, parallel)
                    quarantined = too_deep.copy()
                    _mark_uncached_quarantine(engine, quarantined)

        unavailable |= _cycle_unavailable(
            engine, data, engine_sheets, warnings, evaluation_failed=not engine_alive
        )
        scratch_ready = _ensure_scratch(engine)
    return EngineSession(
        engine,
        engine_sheets,
        engine_alive,
        quarantined,
        scratch_ready,
        unavailable=unavailable,
    )


def _cycle_unavailable(
    engine, data, engine_sheets, warnings, reachable=None, *, evaluation_failed=False
):
    """Keep iteration endpoints from masquerading as converged calculations.

    Telemetry reports convergence per component in aggregate, without cell
    addresses. If any component was capped, all cyclic components and their
    dependents conservatively use file caches. Acyclic branches stay usable.
    """
    try:
        telemetry = engine.last_cycle_telemetry()
    except (AttributeError, RuntimeError):
        return set()
    iterate, _count, _delta = _iteration_options(data)
    capped = telemetry.capped_sccs
    if (
        not evaluation_failed
        and not capped
        and (iterate or not telemetry.live_cycles_witnessed)
    ):
        if iterate and telemetry.iterated_sccs:
            warnings.append(
                "Iterative calculation converged under engine criteria for "
                f"{telemetry.converged_sccs} "
                f"cycle component(s) within {telemetry.max_passes_single_scc} "
                "pass(es), using the workbook's iteration limits. A circular "
                "model can have multiple fixed points; the saved values may "
                "reflect different starting values or evaluation order"
            )
        return set()
    cells = {
        (s, r, c): f
        for s, r, c, f in _iter_formulas_xml(data)
        if s in engine_sheets and (reachable is None or (s, r, c) in reachable)
    }
    labels = {f"{quote_sheet(s)}!{a1(r, c)}": (s, r, c) for s, r, c in cells}
    conservative = False
    try:
        engine.get_eval_plan([])
        trace = engine.trace(
            list(labels),
            direction=fz.TraceDirection.Dependents,
            max_depth=TRACE_MAX_DEPTH,
            max_nodes=TRACE_MAX_NODES,
            max_links=TRACE_MAX_LINKS,
            max_work=TRACE_MAX_WORK,
            range_member_budget=TRACE_RANGE_MEMBERS,
        )
        if trace.truncation.incomplete:
            raise ValueError("cycle trace incomplete")
        adjacency = {key: set() for key in trace.nodes.keys()}
        indegree = dict.fromkeys(adjacency, 0)
        for link in trace.links:
            source = link.source.address
            for target in link.targets:
                destination = target.node.address
                if destination not in adjacency[source]:
                    adjacency[source].add(destination)
                    indegree[destination] += 1
        pending = deque(key for key, degree in indegree.items() if degree == 0)
        while pending:
            for destination in adjacency[pending.popleft()]:
                indegree[destination] -= 1
                if indegree[destination] == 0:
                    pending.append(destination)
        affected = {key for key, degree in indegree.items() if degree > 0}
        if not affected and not capped and not telemetry.live_cycles_witnessed:
            return set()
        # Static traces cannot prove a dynamic address independent of a cycle.
        for label, cell in labels.items():
            formula = re.sub(r'"(?:[^"]|"")*"', '""', cells[cell])
            if re.search(r"\b(?:INDIRECT|OFFSET)\s*\(", formula, re.IGNORECASE):
                affected.add(label)
        pending = deque(affected)
        while pending:
            for destination in adjacency.get(pending.popleft(), ()):
                if destination not in affected:
                    affected.add(destination)
                    pending.append(destination)
        unavailable = set()
        for label in affected:
            rect = parse_ref(label)
            if rect is not None and rect.ncells == 1 and rect.sheet in engine_sheets:
                unavailable.add((rect.sheet, rect.r1, rect.c1))
        if not unavailable:
            raise ValueError("cycle cells could not be located")
    except Exception:
        if evaluation_failed and not capped:
            # Missing-sheet errors can prevent the engine from exposing any
            # trace at all. Recover the static cell graph from the package so
            # one broken reference does not hide every independent formula.
            unavailable = _static_cycle_dependents(data, cells)
            if not unavailable:
                return set()
            conservative = unavailable == set(cells)
        else:
            conservative = True
            unavailable = set(cells)
    reason = (
        f"Iterative calculation did not converge in {telemetry.max_passes_single_scc} "
        f"pass(es) for {capped} cycle component(s)"
        if capped
        else "Circular calculation was detected without workbook iteration enabled"
    )
    if evaluation_failed and not capped:
        reason = (
            "Circular calculation could not be verified after incomplete evaluation"
        )
    scope = (
        "all formula readings" if conservative else "cyclic cells and their dependents"
    )
    warnings.append(
        f"{reason}: {scope} keep file caches where available and omit decomposition; "
        "iteration endpoints are not reported as verified recalculations"
    )
    return unavailable


def _static_cycle_dependents(data, cells):
    from linexcel.resolver import _collect_ref_strings
    from linexcel.structure import read_structure

    names = {
        name.upper(): refs for name, refs in read_structure(data).defined_names.items()
    }
    by_sheet = {}
    for sheet, row, col in cells:
        by_sheet.setdefault(sheet, []).append((row, col))
    for positions in by_sheet.values():
        positions.sort()
    adjacency = {cell: set() for cell in cells}
    indegree = dict.fromkeys(cells, 0)
    work = 0
    for destination, formula in cells.items():
        try:
            refs = _collect_ref_strings(fz.parse(formula).to_dict())
        except Exception:
            continue
        for ref in refs:
            rect = parse_ref(ref, default_sheet=destination[0])
            for region in [rect] if rect is not None else names.get(ref.upper(), []):
                positions = by_sheet.get(region.sheet, [])
                start = bisect_left(positions, (region.r1, region.c1))
                end = bisect_right(positions, (region.r2, region.c2))
                work += end - start
                if work > TRACE_MAX_WORK:
                    return set(cells)
                for row, col in positions[start:end]:
                    if not region.c1 <= col <= region.c2:
                        continue
                    source = (region.sheet, row, col)
                    if destination not in adjacency[source]:
                        adjacency[source].add(destination)
                        indegree[destination] += 1
    pending = deque(cell for cell, degree in indegree.items() if degree == 0)
    while pending:
        for destination in adjacency[pending.popleft()]:
            indegree[destination] -= 1
            if indegree[destination] == 0:
                pending.append(destination)
    affected = {cell for cell, degree in indegree.items() if degree > 0}
    for cell, formula in cells.items():
        code = re.sub(r'"(?:[^"]|"")*"', '""', formula)
        if re.search(r"\b(?:INDIRECT|OFFSET)\s*\(", code, re.IGNORECASE):
            affected.add(cell)
    pending = deque(affected)
    while pending:
        for destination in adjacency[pending.popleft()]:
            if destination not in affected:
                affected.add(destination)
                pending.append(destination)
    return affected


def _uncached_dependents(engine, data: bytes, too_deep: dict, warnings: list[str]):
    """Keep engine limitations distinct from spreadsheet errors caught by guards.

    IFERROR can catch our NImpl marker, but its fallback is not evidence of
    what Excel would compute from the missing input. Suppress readings and
    decompositions throughout the dependent closure, retaining file caches.
    """
    roots = [cell for cell in too_deep if engine.get_value(*cell) is None]
    if not roots:
        return set()
    # Dynamic addresses do not expose all precedents to a static trace. Their
    # independence from an unavailable value cannot be established safely.
    for sheet, row, col, formula in _iter_formulas_xml(data):
        code = re.sub(r'"(?:[^"]|"")*"', '""', formula)
        if re.search(r"\b(?:INDIRECT|OFFSET)\s*\(", code, re.IGNORECASE):
            roots.append((sheet, row, col))
    try:
        # XLSX loading defers construction of the dependency graph. A plan
        # materializes it without evaluating formulas; trace alone refuses it.
        engine.get_eval_plan([])
        trace = engine.trace(
            [f"{quote_sheet(s)}!{a1(r, c)}" for s, r, c in roots],
            direction=fz.TraceDirection.Dependents,
            max_depth=TRACE_MAX_DEPTH,
            max_nodes=TRACE_MAX_NODES,
            max_links=TRACE_MAX_LINKS,
            max_work=TRACE_MAX_WORK,
            range_member_budget=TRACE_RANGE_MEMBERS,
        )
        if trace.truncation.incomplete:
            raise ValueError("dependent trace budget exhausted")
        unavailable = set(roots)
        for key in trace.nodes.keys():
            rect = parse_ref(key)
            if rect is not None and rect.ncells == 1 and rect.sheet is not None:
                unavailable.add((rect.sheet, rect.r1, rect.c1))
        return unavailable
    except Exception:
        warnings.append(
            "The dependents of an uncached over-deep formula could not be fully "
            "traced; formula readings use file caches conservatively because "
            "their independence from the missing input could not be established"
        )
        return {(s, r, c) for s, r, c, _ in _iter_formulas_xml(data)}


def _mark_uncached_quarantine(engine, quarantined: dict) -> None:
    """A missing cached result is unknown, never the numeric value of blank.

    NImpl is an engine error that propagates through arithmetic and ranges,
    and the value resolver already treats it as uncomputed. Cached values
    remain constants so downstream calculations can use the stored reading.
    """
    for sheet, row, col in quarantined:
        if engine.get_value(sheet, row, col) is None:
            engine.set_value(sheet, row, col, {"type": "Error", "kind": "NImpl"})


def _evaluate_targets(
    engine,
    engine_sheets: set[str],
    targets: list[tuple[str, int, int]],
    warnings: list[str],
    progress,
) -> tuple[set[tuple[str, int, int]], bool]:
    """Evaluate only the upstream subgraph of ``targets``; return its cells.

    The global ``evaluate_all`` is never called: the trace walks the declared
    precedents of the target cells, the closure is evaluated in one
    ``evaluate_cells`` batch, and the reachable set comes back so the sweep
    can skip everything else.

    The batch is all-or-nothing like ``evaluate_all``: one broken precedent
    fails it whole. The fallback evaluates the requested targets one at a
    time — the cells the user actually asked for — and leaves the rest to
    the per-cell recovery, which is what a global failure would have got.
    """
    unknown = sorted({sheet for sheet, _, _ in targets} - engine_sheets)
    if unknown:
        raise ValueError(
            f"--target names sheet(s) this workbook does not have: "
            f"{', '.join(unknown)}. It has: {', '.join(sorted(engine_sheets))}."
        )
    progress.step(f"tracing upstream of {len(targets)} target(s)")
    roots = [f"{quote_sheet(sheet)}!{a1(row, col)}" for sheet, row, col in targets]
    trace = engine.trace(
        roots,
        direction=fz.TraceDirection.Precedents,
        max_depth=TRACE_MAX_DEPTH,
        max_nodes=TRACE_MAX_NODES,
        max_links=TRACE_MAX_LINKS,
        max_work=TRACE_MAX_WORK,
        range_member_budget=TRACE_RANGE_MEMBERS,
    )
    reachable: set[tuple[str, int, int]] = set(targets)
    for key in trace.nodes.keys():
        rect = parse_ref(key)
        if rect is not None and rect.ncells == 1 and rect.sheet in engine_sheets:
            reachable.add((rect.sheet, rect.r1, rect.c1))
    if trace.truncation.incomplete:
        warnings.append(
            "The upstream trace of the target cell(s) hit its budget: part of "
            "the subgraph is missing from the lineage. The engine may still "
            "evaluate omitted precedents to compute the requested targets"
        )
    # The evaluation plan counts layers only for cells still dirty, so it is
    # read here, before the evaluation — after it the same plan comes back
    # empty. An indicator of risk, not a duration estimate.
    try:
        layers = len(engine.get_eval_plan(sorted(reachable)).layers)
    except Exception:
        layers = 0  # a plan that will not build says nothing usable
    if layers >= CHAIN_LAYERS_WARNING:
        warnings.append(
            f"Long dependency chains: the targeted subgraph stacks "
            f"{layers:,} evaluation layers. Such workbooks are where the "
            f"step-by-step decomposition hits its time budget (raise it with "
            f"--time-budget) and where an analysis runs long — a risk "
            f"indicator, not a duration estimate"
        )
    progress.step(f"evaluating {len(reachable):,} cell(s)")
    succeeded = True
    try:
        engine.evaluate_cells(sorted(reachable))
    except Exception as exc:
        succeeded = False
        failed = []
        for sheet, row, col in targets:
            try:
                engine.evaluate_cell(sheet, row, col)
            except Exception:
                failed.append(f"{sheet}!{a1(row, col)}")
        note = (
            f"the target(s) themselves did not evaluate: {', '.join(failed)}. "
            if failed
            else "the requested targets were evaluated one by one instead. "
        )
        warnings.append(
            f"Targeted evaluation of {len(reachable):,} cell(s) failed as one "
            f"batch ({exc}); {note}Cells that keep the value stored in the "
            f"file are listed by the per-cell recovery"
        )
    return reachable, succeeded


def _formulas_gone(
    engine,
    data: bytes,
    engine_sheets: set[str],
    quarantined: dict[tuple[str, int, int], str],
) -> bool:
    """Whether the failed evaluation took the formulas down with it.

    A rebuild from the bytes was the unconditional answer to a failed
    ``evaluate_all``, on the grounds that the engine then reports no formula
    at all. On 0.9.3 that is no longer what happens — ``get_formulas`` and
    per-cell ``evaluate_cell`` still answer after the failure (measured on a
    missing-sheet reference) — so the third ``from_bytes`` is probed for
    rather than paid: one formula cell the quarantine did not blank is read
    back, and only an empty answer sends the run through the rebuild.
    """
    for sheet, row, col, _formula in _iter_formulas_xml(data):
        if sheet not in engine_sheets or (sheet, row, col) in quarantined:
            continue
        try:
            return not engine.get_formula(sheet, row, col)
        except Exception:
            return True
    return False


def _find_unresolvable(
    data: bytes, engine_sheets: set[str]
) -> dict[tuple[str, int, int], str]:
    """The formulas that stop ``evaluate_all``, keyed by their cell.

    Only ever called after a global evaluation has already failed, so both ways
    of being wrong are safe: quarantining a formula that would in fact have
    evaluated costs it the same per-cell recovery every cell was getting anyway,
    and missing one simply leaves the retry failing as before.

    The suspects are read from the sheet XML, never from the engine. Reading
    them with ``get_formulas`` would force the engine to build its dependency
    graph one cell at a time — measured quadratic, about 18 s for 20,000 cells,
    i.e. half an hour for the kind of workbook that lands here — and that graph
    build is exactly what the failed ``evaluate_all`` proved impossible. The
    XML scan streams the parts the file already carries and stays linear.
    """
    suspects: dict[tuple[str, int, int], str] = {}
    for sheet, row, col, formula in _iter_formulas_xml(data):
        if sheet in engine_sheets and _is_unresolvable(formula, engine_sheets):
            suspects[(sheet, row, col)] = formula
    return suspects


def _iter_formulas_xml(data: bytes):
    """Yield ``(sheet, row, col, formula)`` for every formula cell in the package.

    The text is the file's own, ``=``-prefixed to match what the engine hands
    back from ``get_formulas``, with XML entities resolved. Shared formulas
    (``<f t="shared" si="n"/>``) yield their master's text untranslated: the
    relative references a translation would shift are exactly the ones no
    suspect check looks at, and the sheet qualifiers and brackets it does look
    at are identical in every cell of the share group.
    """
    from linexcel.loader import _parse_sheet_targets

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
            rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
            for name, zpath in _parse_sheet_targets(wb_xml, rels).items():
                if not zpath.endswith(".xml"):
                    continue
                try:
                    sheet_xml = zf.read(zpath)
                except KeyError:
                    continue
                if b"<f" not in sheet_xml:
                    continue
                shared: dict[bytes, bytes] = {}
                for cell in _CELL_RE.finditer(sheet_xml):
                    f = _F_RE.search(cell.group(3))
                    if f is None:
                        continue
                    text = f.group(2)
                    if text is None:
                        # A shared-formula slave: same formula as its master.
                        si = _SI_RE.search(f.group(1))
                        text = shared.get(si.group(1)) if si else None
                        if text is None:
                            continue
                    elif b't="shared"' in f.group(1):
                        si = _SI_RE.search(f.group(1))
                        if si:
                            shared[si.group(1)] = text
                    formula = "=" + unescape(
                        text.decode("utf-8", "replace"), _XML_ENTITIES
                    )
                    yield (
                        name,
                        int(cell.group(2)),
                        col_to_num(cell.group(1).decode("ascii")),
                        formula,
                    )
    except Exception:
        # A package this best-effort scan cannot read leaves the retry to fail
        # as it did before quarantine existed.
        return


def ast_depth(formula: str) -> int | None:
    """Depth of ``formula``'s parse tree, or ``None`` when it does not parse.

    The walk is iterative: a recursive one would hit Python's own recursion
    limit on exactly the formulas this exists to measure. ``fz.parse`` itself
    holds up at those depths — it is the evaluator, not the parser, that
    overflows.
    """
    try:
        root = fz.parse(formula if formula.startswith("=") else "=" + formula).to_dict()
    except Exception:
        return None
    depth = 0
    stack: list[tuple[object, int]] = [(root, 1)]
    while stack:
        node, level = stack.pop()
        depth = max(depth, level)
        if isinstance(node, dict):
            children = node.values()
        elif isinstance(node, list):
            children = node
        else:
            continue
        stack.extend((v, level + 1) for v in children if isinstance(v, (dict, list)))
    return depth


def is_too_deep(formula: str, max_depth: int = MAX_AST_DEPTH) -> bool:
    """Whether evaluating ``formula`` risks the stack-overflow abort.

    A tree can never nest deeper than the formula is long — every level spends
    at least one character — so the parse is paid only for formulas long
    enough to be dangerous.
    """
    if len(formula) <= max_depth:
        return False
    depth = ast_depth(formula)
    return depth is not None and depth > max_depth


def _find_too_deep(
    data: bytes, engine_sheets: set[str], max_depth: int = MAX_AST_DEPTH
) -> tuple[dict[tuple[str, int, int], str], tuple[str, int, int]]:
    """The cells whose formula is too deep to evaluate safely, keyed by cell.

    Read from the sheet XML like the other quarantine suspects, never from the
    engine — the scan must stay linear and must not ask the engine to build
    the very graph the evaluation is about to choke on. Also returned: the
    first offender, so the warning can name a cell rather than a count.
    """
    suspects: dict[tuple[str, int, int], str] = {}
    first: tuple[str, int, int] | None = None
    for sheet, row, col, formula in _iter_formulas_xml(data):
        if sheet not in engine_sheets or not is_too_deep(formula, max_depth):
            continue
        suspects[(sheet, row, col)] = formula
        if first is None:
            first = (sheet, row, col)
    return suspects, first or ("", 0, 0)


def _without_chartsheets(data: bytes) -> tuple[bytes, list[str]]:
    """The package with its chartsheet ``<sheet>`` entries cut, and their names.

    formualizer's reader refuses a workbook holding a chartsheet outright —
    ``Expecting a worksheet, got chartsheet`` — however many real worksheets
    sit beside it. A chartsheet holds a chart and no cells, so cutting its
    entry from ``workbook.xml`` costs the lineage nothing and lets the rest
    of the workbook through. The names come back so the caller can say which.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    except Exception:
        return data, []
    chartsheet_ids: set[str] = set()
    for rel in re.finditer(r"<Relationship\b([^>]+)>", rels):
        attrs = rel.group(1)
        type_m = re.search(r'\bType="([^"]*)"', attrs)
        id_m = re.search(r'\bId="([^"]+)"', attrs)
        if type_m and id_m and type_m.group(1).endswith("/chartsheet"):
            chartsheet_ids.add(id_m.group(1))
    if not chartsheet_ids:
        return data, []

    names: list[str] = []

    def drop(m: re.Match[str]) -> str:
        attrs = m.group(1)
        rid_m = re.search(r'\br:id="([^"]+)"', attrs)
        if not rid_m or rid_m.group(1) not in chartsheet_ids:
            return m.group(0)
        name_m = re.search(r'\bname="([^"]*)"', attrs)
        names.append(unescape(name_m.group(1), _XML_ENTITIES) if name_m else "?")
        return ""

    # ``<sheet>`` has no children, so the element is always self-closing.
    new_wb_xml = re.sub(r"<sheet\b([^>]+)/>", drop, wb_xml)
    if not names:
        return data, []
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as zf,
        zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for entry in zf.infolist():
            payload = zf.read(entry.filename)
            if entry.filename == "xl/workbook.xml":
                payload = new_wb_xml.encode("utf-8")
            zout.writestr(entry, payload)
    return out.getvalue(), names


def _blank_formulas_in_package(
    data: bytes, quarantined: dict[tuple[str, int, int], str]
) -> bytes:
    """Copy the package with the quarantined cells' ``<f>`` elements cut out.

    Blanking through the engine (``set_value`` / ``set_values_batch``) costs it
    a graph notice per cell — measured near 1.5 ms each even batched, so the
    tens of thousands of external-reference cells a real workbook lands here
    with meant minutes spent mutating. Cutting the elements out of the XML is a
    linear pass, and a cell that keeps its cached ``<v>`` comes back as the
    constant the user last saw, which is exactly the fallback value the
    quarantine used to throw away.

    A shared-formula master carries the text its slaves only name by index, so
    a group is cut whole or not at all; the suspect scan reads every slave as
    its master's text, which guarantees that.
    """
    from linexcel.loader import _parse_sheet_targets

    doomed: dict[str, set[tuple[int, int]]] = {}
    for sheet, row, col in quarantined:
        doomed.setdefault(sheet, set()).add((row, col))

    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
        path_of = _parse_sheet_targets(wb_xml, rels)
        sheet_of_path = {path: name for name, path in path_of.items()}
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for entry in zf.infolist():
                payload = zf.read(entry.filename)
                cells = doomed.get(sheet_of_path.get(entry.filename, ""))
                if cells:

                    def strip(m: re.Match[bytes], cells=cells) -> bytes:
                        key = (int(m.group(2)), col_to_num(m.group(1).decode("ascii")))
                        if key not in cells:
                            return m.group(0)
                        return _F_RE.sub(b"", m.group(0), count=1)

                    payload = _CELL_RE.sub(strip, payload)
                zout.writestr(entry, payload)
    return out.getvalue()


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
