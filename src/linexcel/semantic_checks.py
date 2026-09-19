"""Small runtime capability probes, not an oracle for a user's workbook.

The probes identify known engine limitations. Graph annotations conservatively
identify formulas that may encounter them; they never rewrite values or claim
that an affected formula is necessarily wrong.
"""

from __future__ import annotations

import copy
import datetime
import io
from collections import defaultdict, deque
from functools import lru_cache
from importlib.metadata import version
from typing import Any

import formualizer as fz
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.workbook.defined_name import DefinedName

from linexcel.engine import _open_workbook, is_too_deep

# Expected results use Excel's number-before-text comparison ordering and
# date serial arithmetic. Literal and referenced operands are both exercised.
_CASES = (
    ("number_space_lt", "mixed_type_comparison", '=10<" "', True),
    ("space_number_lt", "mixed_type_comparison", '=" "<10', False),
    ("number_text_eq", "mixed_type_comparison", '="0"=0', False),
    ("reference_space_lt", "mixed_type_comparison", "=B1<B2", True),
    ("reference_numeric_text_eq", "mixed_type_comparison", "=B3=0", False),
    ("space_guard", "mixed_type_comparison", "=IF(B1<B2,7,9)", 7),
    ("ordinary_text_lt", "mixed_type_comparison", '=10<"abc"', True),
    ("text_case", "text_comparison", '="abc"="ABC"', True),
    ("trailing_space", "text_comparison", '="abc"="abc "', False),
    ("numeric_comparison", "numeric_comparison", "=2<3", True),
    ("typed_date_difference", "typed_date_arithmetic", "=A2-A1", 1),
    ("date_serial_subtraction", "typed_date_arithmetic", "=DATE(2026,2,1)-46054", 0),
    ("date_year", "date_functions", "=YEAR(A1)", 2024),
    ("date_difference", "date_functions", "=DATE(2024,1,2)-DATE(2024,1,1)", 1),
    ("legacy_standard_normal", "legacy_normal", "=NORMSDIST(0)", 0.5),
    ("legacy_normal", "legacy_normal", "=NORMDIST(0,0,1,TRUE)", 0.5),
    ("name_numeric_constant", "defined_name_constant", "=GlobalRate*10", 20),
    (
        "name_text_constant",
        "defined_name_constant",
        "=CompanyName",
        "Synthetic Company",
    ),
    ("name_local_constant", "defined_name_constant", "=ScopedRate*10", 30),
    ("name_reference", "defined_name_reference", "=InputAmount+1", 11),
)


def _matches(observed: Any, expected: Any) -> bool:
    if isinstance(expected, bool):
        return isinstance(observed, bool) and observed == expected
    return not isinstance(observed, bool) and observed == expected


def _literal_kind(node: dict[str, Any]) -> str | None:
    if node.get("node_type") != "Literal":
        return None
    value = node.get("value")
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    return None


@lru_cache(maxsize=1)
def _probe_engine_capabilities() -> dict[str, Any]:
    checks = []
    for epoch in ("1900", "1904"):
        workbook = Workbook()
        if epoch == "1904":
            workbook.epoch = CALENDAR_MAC_1904
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Probe"
        sheet["A1"] = datetime.datetime(2024, 1, 1)
        sheet["A2"] = datetime.datetime(2024, 1, 2)
        sheet["B1"], sheet["B2"], sheet["B3"] = 10, " ", "0"
        workbook.defined_names.add(DefinedName("GlobalRate", attr_text="2"))
        workbook.defined_names.add(
            DefinedName("CompanyName", attr_text='"Synthetic Company"')
        )
        workbook.defined_names.add(DefinedName("ScopedRate", attr_text="9"))
        sheet.defined_names.add(DefinedName("ScopedRate", attr_text="3"))
        workbook.defined_names.add(DefinedName("InputAmount", attr_text="Probe!$B$1"))
        for row, (_key, _feature, formula, _expected) in enumerate(_CASES, 1):
            sheet.cell(row, 3, formula)
        output = io.BytesIO()
        workbook.save(output)
        engine = _open_workbook(output.getvalue(), parallel=False)
        engine.evaluate_all()
        for row, (key, feature, formula, expected) in enumerate(_CASES, 1):
            if key == "date_serial_subtraction" and epoch == "1904":
                expected = -1462
            observed = engine.get_value("Probe", row, 3)
            matched = _matches(observed, expected)
            checks.append(
                {
                    "id": f"{key}:{epoch}",
                    "feature": feature,
                    "epoch": epoch,
                    "formula": formula,
                    "expected": expected,
                    "observedType": type(observed).__name__,
                    "observed": observed
                    if isinstance(observed, (dict, str, int, float, bool, type(None)))
                    else {"type": type(observed).__name__, "value": str(observed)},
                    "status": "passed" if matched else "known_mismatch",
                }
            )
    return {
        "engineVersion": version("formualizer"),
        "status": "known_mismatch"
        if any(c["status"] == "known_mismatch" for c in checks)
        else "probes_passed",
        "scope": "Synthetic capability probes only; passing does not verify "
        "workbook values",
        "checks": checks,
        "inputs": {
            "A1": {"type": "datetime", "value": "2024-01-01T00:00:00"},
            "A2": {"type": "datetime", "value": "2024-01-02T00:00:00"},
            "B1": 10,
            "B2": " ",
            "B3": "0",
        },
        "definedNames": [
            {"name": "GlobalRate", "scope": "workbook", "expression": "2"},
            {
                "name": "CompanyName",
                "scope": "workbook",
                "expression": '"Synthetic Company"',
            },
            {"name": "ScopedRate", "scope": "workbook", "expression": "9"},
            {
                "name": "ScopedRate",
                "scope": "sheet",
                "scope_sheet": "Probe",
                "expression": "3",
            },
            {
                "name": "InputAmount",
                "scope": "workbook",
                "expression": "Probe!$B$1",
            },
        ],
    }


def engine_capabilities() -> dict[str, Any]:
    """Fresh report object; native probes execute once per worker process."""
    try:
        return copy.deepcopy(_probe_engine_capabilities())
    except Exception as exc:
        return {"status": "unavailable", "checks": [], "reason": str(exc)}


def _constant_name_references(
    ast: dict[str, Any], evidence: dict[str, Any], sheet: str | None
) -> bool:
    """Identify proven numeric/text constants with source and lexical scope.

    Partial metadata cannot rule out an omitted sheet-local definition that
    shadows a recorded global name. Do not infer support from missing evidence.
    """
    if evidence.get("status") != "complete":
        return False
    definitions = {
        (
            (item.get("scope_sheet") or "").casefold()
            if item.get("scope") == "sheet"
            else "",
            (item.get("name") or "").casefold(),
        ): item
        for item in evidence.get("definitions", [])
        if item.get("scope") in {"sheet", "workbook"}
    }
    sheets = {name.casefold() for name in evidence.get("sheets", []) if name}
    pending = [(ast, frozenset())]
    while pending:
        part, bound = pending.pop()
        if not isinstance(part, dict):
            continue
        if part.get("node_type") == "Reference":
            reference = part.get("reference", "")
            qualifier, separator, name = reference.rpartition("!")
            if not separator:
                name = reference
                if name.casefold() in bound:
                    continue
                active_sheet = (sheet or "").casefold()
            else:
                if qualifier.startswith("'") and qualifier.endswith("'"):
                    qualifier = qualifier[1:-1].replace("''", "'")
                active_sheet = qualifier.casefold()
                if active_sheet not in sheets:
                    continue
            definition = definitions.get((active_sheet, name.casefold()))
            if definition is None:
                definition = definitions.get(("", name.casefold()))
            expression = definition.get("expression") if definition else None
            if isinstance(expression, str) and not is_too_deep(expression):
                try:
                    parsed = fz.parse(
                        expression if expression.startswith("=") else "=" + expression
                    ).to_dict()
                except Exception:
                    continue
                if _literal_kind(parsed) in {"number", "text"}:
                    return True
                if (
                    parsed.get("node_type") == "UnaryOp"
                    and parsed.get("operator") in {"+", "-"}
                    and _literal_kind(parsed.get("operand", {})) == "number"
                ):
                    return True
        if part.get("node_type") == "Function":
            name = str(part.get("name", "")).upper().removeprefix("_XLFN.")
            args = part.get("args", [])
            if name == "LET" and len(args) >= 3 and len(args) % 2 == 1:
                scoped = bound
                for index in range(0, len(args) - 1, 2):
                    pending.append((args[index + 1], scoped))
                    binding = args[index].get("reference", "")
                    scoped = scoped | {binding.casefold()}
                pending.append((args[-1], scoped))
                continue
            if name == "LAMBDA" and args:
                scoped = bound | {
                    arg.get("reference", "").casefold() for arg in args[:-1]
                }
                pending.append((args[-1], scoped))
                continue
        for value in part.values():
            if isinstance(value, dict):
                pending.append((value, bound))
            elif isinstance(value, list):
                pending.extend(
                    (item, bound) for item in value if isinstance(item, dict)
                )
    return False


def annotate_semantic_risks(
    graph: dict[str, Any], *, capabilities: dict[str, Any] | None = None
) -> None:
    """Attach bounded, conservative risk annotations without changing values."""
    report = (
        engine_capabilities() if capabilities is None else copy.deepcopy(capabilities)
    )
    meta = graph.setdefault("meta", {})
    meta["engineCapabilities"] = report
    failures = {
        check["feature"]
        for check in report.get("checks", [])
        if check["status"] == "known_mismatch"
    }
    if not failures:
        if report.get("status") == "unavailable":
            meta.setdefault("warnings", []).append(
                "Engine semantic capability probes were unavailable; "
                "no capability-based verification was possible"
            )
        return
    nodes = {node["id"]: node for node in graph.get("nodes", [])}
    forward: dict[str, set[str]] = defaultdict(set)
    incoming: dict[str, set[str]] = defaultdict(set)
    for edge in graph.get("edges", []):
        if edge.get("kind") == "dep":
            forward[edge["source"]].add(edge["target"])
            incoming[edge["target"]].add(edge["source"])
    affected: dict[str, set[str]] = defaultdict(set)
    for key, node in nodes.items():
        formula = node.get("formula")
        if not formula or is_too_deep(formula):
            continue
        try:
            ast = fz.parse(formula).to_dict()
        except Exception:
            continue
        if "defined_name_constant" in failures and _constant_name_references(
            ast, meta.get("definedNameEvidence", {}), node.get("sheet")
        ):
            affected[key].add("defined_name_constant")
        stack = [ast]
        has_arithmetic = False
        has_date_function = False
        while stack:
            part = stack.pop()
            if not isinstance(part, dict):
                continue
            operator = (
                part.get("operator") if part.get("node_type") == "BinaryOp" else None
            )
            if operator in {"=", "<>", "<", ">", "<=", ">="}:
                left = _literal_kind(part.get("left", {}))
                right = _literal_kind(part.get("right", {}))
                for feature in failures & {
                    "mixed_type_comparison",
                    "text_comparison",
                    "numeric_comparison",
                }:
                    # Both literal operand types are proven. References and
                    # computed operands remain conservative because group
                    # samples cannot establish every cell's runtime type.
                    if (
                        feature == "mixed_type_comparison"
                        and left is not None
                        and left == right
                    ):
                        continue
                    if feature == "text_comparison" and left == right == "number":
                        continue
                    if feature == "numeric_comparison" and left == right == "text":
                        continue
                    affected[key].add(feature)
            if operator in {"+", "-", "*", "/"} and "typed_date_arithmetic" in failures:
                has_arithmetic = True
                if node.get("valueDate") or any(
                    nodes.get(source, {}).get("valueDate") for source in incoming[key]
                ):
                    affected[key].add("typed_date_arithmetic")
            name = str(part.get("name", "")).upper()
            if part.get("node_type") == "Function":
                if name in {"DATE", "DATEVALUE", "EDATE", "EOMONTH"}:
                    has_date_function = True
                if name in {"NORMSDIST", "NORMDIST"} and "legacy_normal" in failures:
                    affected[key].add("legacy_normal")
                if name in {"DATE", "YEAR"} and "date_functions" in failures:
                    affected[key].add("date_functions")
            for value in part.values():
                if isinstance(value, dict):
                    stack.append(value)
                elif isinstance(value, list):
                    stack.extend(item for item in value if isinstance(item, dict))
        if has_arithmetic and has_date_function:
            affected[key].add("typed_date_arithmetic")
    roots = {key: set(reasons) for key, reasons in affected.items()}
    queue = deque(roots)
    while queue:
        source = queue.popleft()
        for target in forward[source]:
            previous = len(affected[target])
            affected[target].update(affected[source])
            if len(affected[target]) != previous:
                queue.append(target)
    for key, reasons in affected.items():
        if key not in nodes:
            continue
        nodes[key]["verification"] = "unverified_engine_semantics"
        nodes[key]["semanticRisks"] = [
            {
                "feature": reason,
                "origin": "formula"
                if reason in roots.get(key, set())
                else "dependency",
                "status": "potentially_affected",
            }
            for reason in sorted(reasons)
        ]
    meta.setdefault("warnings", []).append(
        "Engine capability probes detected known semantic mismatches: "
        + ", ".join(sorted(failures))
        + f". {len(affected)} formula/dependent node(s) are conservatively marked "
        "unverified; this does not establish that every marked value is wrong. "
        "Original values and their provenance are preserved"
    )
