"""Deterministic lazy workbook graph and exact, sparse target closure.

These entry points must be called in a resource-limited child process. Import
never initializes or evaluates the native formula engine. Unsupported dependencies fail
closed before any evaluation; no clipped ranges or cached formula substitutions.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import math
import re
from collections import defaultdict
from dataclasses import replace
from pathlib import Path
from typing import Any

from openpyxl.formula import Tokenizer
from openpyxl.utils.datetime import to_excel

from linexcel.refs import (
    a1,
    parse_ref,
    parse_ref_detailed,
    quote_sheet,
    ref_to_r1c1,
    split_sheet_prefix,
)
from linexcel.values import (
    ERROR_KIND_TEXT,
    EXCEL_ERRORS,
    _excel_error_text,
    _is_uncomputed,
)

_INPUT_ERROR_KINDS = {value: kind for kind, value in ERROR_KIND_TEXT.items()}

_VOLATILE = {"RAND", "RANDBETWEEN", "RANDARRAY", "NOW", "TODAY"}
_PATTERN_CELL = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*"
_PATTERN_REFERENCE = re.compile(
    rf"(?:{_PATTERN_CELL}(?::{_PATTERN_CELL})?|"
    r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)"
)

_DYNAMIC = {
    "INDIRECT",
    "OFFSET",
    "CELL",
    "INFO",
    "SHEET",
    "SHEETS",
    "SUBTOTAL",
    "AGGREGATE",
    "FORMULATEXT",
    "RTD",
    "WEBSERVICE",
    "STOCKHISTORY",
}


def _json(value: Any) -> Any:
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _engine_value(value: Any) -> Any:
    """Normalize real Excel results; never stringify a native limitation."""
    if _is_uncomputed(value):
        raise ValueError(f"Engine could not calculate this value: {value.get('kind')}")
    error = _excel_error_text(value)
    if error is not None:
        return error
    if isinstance(value, (list, tuple)):
        return [_engine_value(item) for item in value]
    if value is None or isinstance(value, (str, int, float, bool, dt.date, dt.time)):
        return _json(value)
    raise ValueError(f"Unsupported engine value representation: {type(value).__name__}")


def _id(sheet: str, coordinate: str) -> str:
    return f"{quote_sheet(sheet)}!{coordinate}"


def _pattern_signature(
    node: dict, nodes: dict[str, dict]
) -> tuple[str, list, str | None]:
    """Engine-free R1C1 key; preserve names, literals and operator token types."""
    formula = node["formula"]
    if node.get("diagnostics"):
        return (
            formula,
            [],
            "Formula has unsupported or ambiguous structural diagnostics.",
        )
    try:
        tokens = Tokenizer(formula).items
        identity = []
        readable = []
        for token in tokens:
            value = token.value
            tag = (token.type, token.subtype)
            if token.type == "OPERAND" and token.subtype == "RANGE":
                binding = nodes.get(node.get("reference_bindings", {}).get(value), {})
                if binding.get("kind") == "name":
                    # A name such as TVA or RC is not a column/cell reference.
                    # Include its resolved scope in the identity; never rewrite it.
                    tag = ("DEFINED_NAME", binding["id"])
                    value = binding["name"]
                else:
                    detail = parse_ref_detailed(value)
                    _, body = split_sheet_prefix(value)
                    if (
                        "[" in value
                        or detail is None
                        or not _PATTERN_REFERENCE.fullmatch(body)
                        or binding.get("kind") == "unresolved"
                    ):
                        return (
                            formula,
                            [],
                            f"Reference cannot be canonicalized safely: {value}",
                        )
                    converted = ref_to_r1c1(value, node["row"], node["column"])
                    if converted is None:
                        return (
                            formula,
                            [],
                            f"Reference cannot be canonicalized safely: {value}",
                        )
                    value = converted
                    tag = ("REFERENCE", "R1C1")
            # Do not fold string literals or remove whitespace: a space may be
            # Excel's intersection operator. This prefers false negatives to
            # grouping formulas whose syntax has different meaning.
            identity.append([*tag, value])
            readable.append(value)
        return "=" + "".join(readable), identity, None
    except (ValueError, IndexError) as exc:
        return formula, [], f"Formula cannot be canonicalized safely: {exc}"


def _member_ranges(members: list[dict]) -> list[str]:
    """Tile actual members with exact rectangles; never fill gaps in a group."""
    rows: dict[int, list[int]] = defaultdict(list)
    for node in members:
        rows[node["row"]].append(node["column"])
    rectangles = []
    active: dict[tuple[int, int], tuple[int, int]] = {}
    for row, columns in sorted(rows.items()):
        columns.sort()
        spans = []
        start = end = columns[0]
        for column in columns[1:]:
            if column == end + 1:
                end = column
            else:
                spans.append((start, end))
                start = end = column
        spans.append((start, end))
        next_active = {}
        for span in spans:
            previous = active.pop(span, None)
            if previous and previous[1] == row - 1:
                next_active[span] = (previous[0], row)
            else:
                if previous:
                    rectangles.append((previous[0], span[0], previous[1], span[1]))
                next_active[span] = (row, row)
        for span, (r1, r2) in active.items():
            rectangles.append((r1, span[0], r2, span[1]))
        active = next_active
    for span, (r1, r2) in active.items():
        rectangles.append((r1, span[0], r2, span[1]))
    return [
        a1(r1, c1) if (r1, c1) == (r2, c2) else f"{a1(r1, c1)}:{a1(r2, c2)}"
        for r1, c1, r2, c2 in sorted(rectangles)
    ]


def build_structure(path: Path, refs_dir: Path | None = None) -> dict:
    """Explicit complete export for library callers, built from the disk index.

    The interactive server uses ``build_index`` and bounded pages instead.
    """
    from tempfile import TemporaryDirectory

    from linexcel.lazy_store import build_index, export_structure

    with TemporaryDirectory(prefix="linexcel-export-") as directory:
        index_path = Path(directory) / "index.sqlite"
        build_index(path, index_path, refs_dir)
        return export_structure(index_path)


def _absolute_name(node: dict) -> str | None:
    reference = node["formula"][1:]
    if "[" in reference:
        return None
    rect = parse_ref(reference, default_sheet=node.get("sheet"))
    if rect is None or rect.sheet is None:
        return None
    body = reference.rsplit("!", 1)[-1]
    absolute = r"(?:\$[A-Za-z]{1,3}\$[0-9]+|\$[A-Za-z]{1,3}|\$[0-9]+)"
    if not re.fullmatch(absolute + r"(?::" + absolute + r")?", body):
        return None
    return rect.to_a1()


def _name_literal(node: dict) -> tuple[str, Any, str] | None:
    """A defined scalar literal, never a formula or relative reference.

    Parentheses keep a signed number or error a single operand when substituted
    into a consuming formula. The source name/formula remains in the graph.
    """
    body = node["formula"][1:].strip()
    if re.fullmatch(r'"(?:[^"]|"")*"', body):
        return f"({body})", body[1:-1].replace('""', '"'), "scalar"
    if body.upper() in {"TRUE", "FALSE"}:
        return f"({body.upper()})", body.upper() == "TRUE", "scalar"
    if body in EXCEL_ERRORS:
        return f"({body})", body, "error"
    if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", body):
        value = float(body)
        if math.isfinite(value):
            return f"({body})", value, "scalar"
    return None


def _name_expression(node: dict) -> str | None:
    reference = _absolute_name(node)
    if reference is not None:
        return reference
    literal = _name_literal(node)
    return literal[0] if literal is not None else None


def _cache_comparison(
    node: dict, raw: Any, normalized: Any, epoch: dt.datetime
) -> dict:
    comparison_value = (
        to_excel(raw, epoch)
        if isinstance(raw, (dt.datetime, dt.date, dt.time))
        else normalized
    )
    comparison = "unknown"
    if node.get("cache_status") == "saved":
        saved = node.get("cached_comparison_value", node.get("cached_value"))
        same_boolean_type = isinstance(saved, bool) == isinstance(
            comparison_value, bool
        )
        same_error_type = (node.get("cachedValueKind") == "error") == (
            _excel_error_text(raw) is not None
        )
        comparison = (
            "equal"
            if same_boolean_type and same_error_type and saved == comparison_value
            else "different"
        )
    return {"comparisonValue": comparison_value, "comparison": comparison}


def evaluate_node(
    path: Path,
    node_id: str,
    refs_dir: Path | None = None,
    *,
    index_path: Path | None = None,
    max_closure_nodes: int = 20000,
) -> dict:
    """Evaluate only the exact requested dependency closure in an isolated worker.

    A prebuilt server index avoids reopening or rescanning the workbook. The
    The explicit-node closure bound excludes literal range members, which stay
    on disk until streamed to the native engine. No truncated closure is ever
    calculated; presentation can be bounded independently of source coverage.
    """
    from tempfile import TemporaryDirectory

    from linexcel.lazy_store import WorkbookIndex, build_index

    if index_path is None:
        with TemporaryDirectory(prefix="linexcel-evaluate-") as directory:
            temporary = Path(directory) / "index.sqlite"
            build_index(path, temporary, refs_dir)
            return evaluate_node(
                path,
                node_id,
                refs_dir,
                index_path=temporary,
                max_closure_nodes=max_closure_nodes,
            )
    with WorkbookIndex(index_path) as nodes:
        return _evaluate_index(nodes, node_id, max_closure_nodes)


def _evaluate_index(nodes, node_id: str, max_closure_nodes: int) -> dict:
    from itertools import islice

    if node_id not in nodes:
        raise ValueError(f"Unknown node: {node_id}")
    target = nodes[node_id]
    result: dict[str, Any] = {
        "node_id": node_id,
        "nodeId": node_id,
        "patternId": target.get("patternId"),
        "status": "unsupported",
        "value": None,
        "cachedValue": target.get("cached_value"),
        "cached_value": target.get("cached_value"),
        "comparison": "unknown",
        "provenance": None,
        "steps": [],
        "diagnostics": [],
    }
    if target["kind"] not in {"cell", "input"}:
        result["diagnostics"] = ["Select a single cell for recalculation."]
        return result
    closure: set[str] = set()
    nodes.prepare_input_selection()
    pending = [node_id]
    reachable_deps: dict[str, list[str]] = {}
    while pending:
        nid = pending.pop()
        if nid in closure:
            continue
        if len(closure) >= max_closure_nodes:
            result["diagnostics"] = [
                f"Exact dependency closure exceeds the {max_closure_nodes:,}-node "
                "calculation budget. No partial calculation was performed."
            ]
            result["dependency_count"] = len(closure)
            result["coverage"] = {"status": "not_evaluated", "closureComplete": False}
            return result
        closure.add(nid)
        node = nodes[nid]
        if node["diagnostics"]:
            result["diagnostics"].extend(
                f"{nid}: {message}" for message in node["diagnostics"]
            )
        reachable_deps[nid] = list(node["dependencies"])
        if node["kind"] == "range":
            nodes.select_range_inputs(node["sheet"], node["bounds"])
            members = list(
                islice(
                    nodes.range_members(node["sheet"], node["bounds"], kind="cell"),
                    max_closure_nodes + 1,
                )
            )
            if len(members) > max_closure_nodes:
                result["diagnostics"] = [
                    f"Range {nid} exceeds the {max_closure_nodes:,}-formula "
                    "calculation "
                    "budget. No partial calculation was performed."
                ]
                result["coverage"] = {
                    "status": "not_evaluated",
                    "closureComplete": False,
                }
                return result
            pending.extend(members)
            reachable_deps[nid].extend(members)
        elif node["kind"] == "input":
            nodes.select_input(nid)
        elif node["kind"] == "name":
            if _name_expression(node) is None:
                result["diagnostics"].append(
                    f"{nid}: Only absolute reference or scalar literal defined names "
                    "support recalculation."
                )
        pending.extend(node["dependencies"])
    result["volatile"] = any(nodes[nid].get("volatile", False) for nid in closure)
    input_count = nodes.selected_input_count()
    formula_ids = [nid for nid in sorted(closure) if nodes[nid]["kind"] == "cell"]
    structural_count = sum(
        nodes[nid]["kind"] not in {"cell", "input"} for nid in closure
    )
    total_nodes = len(formula_ids) + structural_count + input_count
    result["dependency_count"] = total_nodes
    # Source inputs stay in the disk set even for million-cell ranges. Check
    # every selected input before starting an engine; an unreadable value must
    # never silently become a blank or let a partial calculation proceed.
    for node in nodes.selected_inputs():
        if node["id"] in closure:
            continue
        for message in node.get("diagnostics", []):
            if len(result["diagnostics"]) < 100:
                result["diagnostics"].append(f"{node['id']}: {message}")
            else:
                result["omittedDiagnostics"] = result.get("omittedDiagnostics", 0) + 1
    # Topological removal detects cycles through compact ranges too. Never
    # pretend to reproduce Excel iterative-calculation settings in a new engine.
    indegrees = {nid: 0 for nid in closure}
    for deps in reachable_deps.values():
        for dep in deps:
            indegrees[dep] += 1
    ready = [nid for nid, degree in indegrees.items() if degree == 0]
    visited = 0
    while ready:
        nid = ready.pop()
        visited += 1
        for dep in reachable_deps[nid]:
            indegrees[dep] -= 1
            if indegrees[dep] == 0:
                ready.append(dep)
    if visited != len(closure):
        result["diagnostics"].append(
            "Circular dependencies require unsupported iteration semantics."
        )

    def make_step(nid):
        node = nodes[nid]
        dependencies = reachable_deps.get(nid, [])
        omitted = 0
        if node["kind"] == "range":
            members = list(
                islice(nodes.range_members(node["sheet"], node["bounds"]), 201)
            )
            dependencies = members[:200]
            if len(members) > 200:
                r1, c1, r2, c2 = node["bounds"]
                count = nodes.db.execute(
                    "SELECT count(*) FROM nodes WHERE sheet=? "
                    "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ? "
                    "AND kind IN ('cell','input')",
                    (node["sheet"], r1, r2, c1, c2),
                ).fetchone()[0]
                omitted = count - len(dependencies)
        step = {
            "node_id": nid,
            "nodeId": nid,
            "label": node.get("label", nid),
            "formula": node.get("formula"),
            "cachedValue": node.get("cachedValue"),
            "cachedValueKind": node.get("cachedValueKind", "scalar"),
            "valueSource": node.get("valueSource", "unknown"),
            "numberFormat": node.get("numberFormat"),
            "dependencies": dependencies,
            "kind": node["kind"],
            "comparison": "unknown",
            "evaluationStatus": (
                "not_evaluated"
                if node["kind"] == "cell"
                else "input"
                if node["kind"] == "input"
                else "structural"
            ),
            "diagnostics": list(node.get("diagnostics", [])),
        }
        if omitted:
            step["dependenciesOmitted"] = omitted
        return step

    # A bounded presentation is independent from exact calculation coverage.
    # Always include the target; keep small closures complete for inspection.
    step_ids = list(dict.fromkeys([node_id, *sorted(closure)]))[:200]
    for nid in nodes.selected_input_sample(200):
        if len(step_ids) >= 200:
            break
        if nid not in step_ids:
            step_ids.append(nid)
    steps: list[dict[str, Any]] = [make_step(nid) for nid in sorted(step_ids)]
    result["steps"] = steps
    result["stepTotal"] = total_nodes
    result["omittedSteps"] = total_nodes - len(steps)
    result["coverage"] = {
        "status": "not_evaluated",
        "totalNodes": total_nodes,
        "formulaCells": len(formula_ids),
        "calculatedFormulaCells": 0,
        "unavailableFormulaCells": 0,
        "unsupportedFormulaCells": 0,
        "notEvaluatedFormulaCells": len(formula_ids),
        "inputCells": input_count,
        "structuralNodes": structural_count,
        "closureComplete": True,
    }
    result["step_semantics"] = (
        "Static dependency closure, including the target and sparse range members. "
        "Formula values are read from engine memory after one target evaluation, "
        "without evaluating steps again or substituting saved formula caches. "
        "The engine may calculate formulas from unselected IF branches; this is "
        "not an execution trace or proof that every value was used by the target. "
        "Inputs are source values; ranges and names are structural. A missing "
        "engine value is unavailable, not zero or the saved cache."
        " Large closures show at most 200 steps and 200 members per range; "
        "omittedSteps/dependenciesOmitted describe presentation omissions only. "
        "Every selected literal input is streamed to the engine."
    )
    if result["diagnostics"]:
        return result

    if target["kind"] == "input":
        # Literal inputs already are deterministic source values. Reading one
        # needs neither a formula engine nor conversion of an error to text.
        value = target.get("value")
        result.update(
            status="completed",
            value=value,
            valueKind=target.get("cachedValueKind", "scalar"),
            comparison="equal",
            provenance="input_read",
        )
        result["coverage"]["status"] = "complete"
        return result

    # Native parser and engine deliberately imported only after closure checks.
    import formualizer as fz

    from linexcel.engine import _register_legacy_normal_functions, is_too_deep
    from linexcel.excel_compat import register_excel_functions, rewrite_formula

    for nid in closure:
        formula = nodes[nid].get("formula")
        if formula and is_too_deep(formula):
            result["diagnostics"].append(
                f"{nid}: Formula exceeds the native engine's safe complexity limit."
            )
    if result["diagnostics"]:
        return result

    from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

    epoch = (
        CALENDAR_MAC_1904
        if nodes.meta["dateSystem"] == "1904"
        else CALENDAR_WINDOWS_1900
    )
    config = fz.EvaluationConfig()
    config.enable_parallel = False
    config.date_system = "1904" if epoch.year == 1904 else "1900"
    engine = fz.Workbook(config=fz.WorkbookConfig(eval_config=config))
    register_excel_functions(engine, config.date_system)
    _register_legacy_normal_functions(engine, config.date_system)
    for sheet in nodes.meta["sheets"]:
        engine.add_sheet(sheet)
    external_sheets = {}
    occupied = set(nodes.meta["sheets"])
    for nid in sorted(closure):
        node = nodes[nid]
        if node.get("external") and node["sheet"] not in external_sheets:
            name = (
                "_linexcel_" + hashlib.sha256(node["sheet"].encode()).hexdigest()[:16]
            )
            while name in occupied:
                name += "_"
            occupied.add(name)
            engine.add_sheet(name)
            external_sheets[node["sheet"]] = name
    for nid in sorted(closure):
        node = nodes[nid]
        if node["kind"] != "cell":
            continue
        sheet = external_sheets.get(node["sheet"], node["sheet"])
        row, column = node["row"], node["column"]
        if node.get("formula"):
            tokens = Tokenizer(node["formula"]).items
            for token in tokens:
                if token.type == "OPERAND" and token.subtype == "RANGE":
                    dep = node.get("reference_bindings", {}).get(token.value)
                    named = nodes.get(dep, {})
                    if named.get("external"):
                        token.value = _id(
                            external_sheets[named["sheet"]], named["cell"]
                        )
                    elif named.get("kind") == "name":
                        replacement = _name_expression(named)
                        assert replacement is not None
                        token.value = replacement
                    elif named.get("sheet"):
                        # Canonical source sheet spelling also fixes native
                        # case-sensitive resolution of Excel references.
                        rect = parse_ref(token.value, default_sheet=sheet)
                        if rect is not None:
                            token.value = replace(rect, sheet=named["sheet"]).to_a1()
            engine.set_formula(
                sheet,
                row,
                column,
                rewrite_formula("=" + "".join(t.value for t in tokens), sheet),
            )
    for node in nodes.selected_inputs():
        sheet = external_sheets.get(node["sheet"], node["sheet"])
        value = node.get("cached_comparison_value", node.get("value"))
        if node.get("sourceDataType") == "e":
            value = {"type": "Error", "kind": _INPUT_ERROR_KINDS[value]}
        if value is not None:
            engine.set_value(sheet, node["row"], node["column"], value)
    value = engine.evaluate_cell(
        external_sheets.get(target["sheet"], target["sheet"]),
        target["row"],
        target["column"],
    )
    # get_value reads the current engine cache; it does not open another
    # evaluation request. In particular, never call evaluate_cell per step:
    # that would recalculate volatile functions and break snapshot coherence.
    presented = {step["nodeId"]: step for step in steps}
    for nid in formula_ids:
        step = presented.get(nid)
        if step is None:
            step = make_step(nid)
        node = nodes[step["nodeId"]]
        result["coverage"]["notEvaluatedFormulaCells"] -= 1
        try:
            raw = (
                value
                if step["nodeId"] == node_id
                else engine.get_value(
                    external_sheets.get(node["sheet"], node["sheet"]),
                    node["row"],
                    node["column"],
                )
            )
        except (ValueError, RuntimeError) as exc:
            step["evaluationStatus"] = "unavailable"
            step["diagnostics"].append(
                "Could not read this operation's engine value: "
                f"{type(exc).__name__}: {exc}"
            )
            result["coverage"]["unavailableFormulaCells"] += 1
            continue
        if raw is None and step["nodeId"] != node_id:
            step["evaluationStatus"] = "unavailable"
            step["diagnostics"].append(
                "The engine did not expose a value after this target evaluation. "
                "A blank result cannot be distinguished from an uncomputed cell."
            )
            result["coverage"]["unavailableFormulaCells"] += 1
            continue
        try:
            calculated = _engine_value(raw)
        except ValueError as exc:
            step["evaluationStatus"] = "unsupported"
            step["diagnostics"].append(str(exc))
            result["coverage"]["unsupportedFormulaCells"] += 1
            continue
        step.update(
            evaluationStatus="calculated",
            calculatedValue=calculated,
            calculatedValueKind=(
                "error"
                if _excel_error_text(raw) is not None
                else "array"
                if isinstance(calculated, list)
                else "scalar"
            ),
            calculatedProvenance="targeted_recalculation",
        )
        step.update(_cache_comparison(node, raw, calculated, epoch))
        result["coverage"]["calculatedFormulaCells"] += 1
    result["coverage"]["status"] = (
        "complete"
        if result["coverage"]["calculatedFormulaCells"]
        == result["coverage"]["formulaCells"]
        else "partial"
    )
    try:
        normalized = _engine_value(value)
    except ValueError as exc:
        result["diagnostics"].append(str(exc))
        return result
    result.update(
        status="completed", value=normalized, provenance="targeted_recalculation"
    )
    result.update(_cache_comparison(target, value, normalized, epoch))
    result["valueKind"] = (
        "error"
        if _excel_error_text(value) is not None
        else "array"
        if isinstance(normalized, list)
        else "scalar"
    )
    return result
