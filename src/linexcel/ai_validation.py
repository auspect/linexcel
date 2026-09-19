"""Bounded quotation checks, never a factual certification of generated prose.

Only Markdown code quotations of complete Excel formulas/function expressions
are compared with the exact source dossier sent to the model. A missing match
is a qualification, not proof of an error: examples, equivalent expressions,
and omitted source evidence can all legitimately fail this lexical check.
"""

from __future__ import annotations

import re
from typing import Any

MAX_TEXT_CHARS = 200_000
MAX_FORMULA_CHARS = 8192
MAX_QUOTES = 256
_FUNCTION = re.compile(r"[A-Za-z_][A-Za-z_0-9.]*\(")
_CODE = re.compile(r"(?<!`)(`{1,3})(?!`)(.*?)\1(?!`)", re.DOTALL)


def _incomplete(expression: str) -> bool:
    # Strings and quoted sheet names must not contribute punctuation.
    plain = re.sub(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'', "", expression)
    return (
        "…" in plain
        or "..." in plain
        or '"' in plain
        or "'" in plain
        or plain.count("(") != plain.count(")")
        or plain.count("{") != plain.count("}")
        or plain.count("[") != plain.count("]")
    )


def _has_reference(expression: str) -> bool:
    """Conservative lexical candidates, not resolved cell/name dependencies."""
    plain = re.sub(r'"(?:[^"]|"")*"', "", expression)
    plain = re.sub(r"[A-Za-z_][A-Za-z_0-9.]*\s*\(", "(", plain)
    plain = re.sub(r"\b(?:TRUE|FALSE)\b", "", plain, flags=re.I)
    # Remove scientific-notation exponents before looking for a reference/name.
    plain = re.sub(r"\d+(?:\.\d+)?[Ee][+-]?\d+", "0", plain)
    return bool(re.search(r"[A-Za-z_$]", plain))


def _canonical(expression: str) -> str:
    """Preserve literals, bracket contents and reference-intersection spaces."""
    expression = expression.strip().removeprefix("=").strip()
    result = []
    index = 0
    while index < len(expression):
        char = expression[index]
        if char in "\"'[":
            end = _protected_end(expression, index)
            result.append(expression[index:end])
            index = end
            continue
        if char.isspace():
            end = index + 1
            while end < len(expression) and expression[end].isspace():
                end += 1
            before = expression[index - 1] if index else ""
            after = expression[end] if end < len(expression) else ""
            # Keep possible intersection operators, including quoted sheets.
            if (before.isalnum() or before in "_$).]'\"") and (
                after.isalnum() or after in "_$(['\""
            ):
                result.append(" ")
            index = end
            continue
        else:
            result.append(char.casefold())
        index += 1
    return "".join(result)


def _protected_end(expression: str, start: int) -> int:
    """Skip a literal, quoted sheet, or nested bracket reference verbatim."""
    delimiter = expression[start]
    depth = 1
    index = start + 1
    while index < len(expression):
        char = expression[index]
        if delimiter == "[":
            if char == "'" and index + 1 < len(expression):
                index += 2  # Structured-reference escape; preserve both characters.
                continue
            if char == "[":
                depth += 1
            elif char == "]":
                depth -= 1
                if depth == 0:
                    return index + 1
        elif char == delimiter:
            if index + 1 < len(expression) and expression[index + 1] == delimiter:
                index += 2
                continue
            return index + 1
        index += 1
    return len(expression)


def _expressions(formula: str) -> list[str]:
    """Collect balanced function subexpressions without invoking a native parser."""
    expressions = [formula]
    stack: list[tuple[int | None, int]] = []
    index = 0
    while index < len(formula):
        char = formula[index]
        if char in "\"'[":
            index = _protected_end(formula, index)
            continue
        elif char == "(":
            prefix = re.search(r"[A-Za-z_][A-Za-z_0-9.]*\s*$", formula[:index])
            stack.append((prefix.start() if prefix else None, index + 1))
        elif char == "," and stack:
            start, argument_start = stack[-1]
            if start is not None:
                expressions.append(formula[argument_start:index])
            stack[-1] = start, index + 1
        elif char == ")" and stack:
            start, argument_start = stack.pop()
            if start is not None:
                expressions.append(formula[argument_start:index])
                expressions.append(formula[start : index + 1])
        index += 1
    return expressions


def _source_formulas(dossier: dict) -> list[tuple[str, str, str]]:
    # Deliberately exclude decomposition, comments, presentation text and values:
    # Formula fields and explicitly supplied deterministic R1C1 forms establish
    # lexical support. Reports distinguish the latter from source formulas.
    records = [dossier]
    for key in ("precedents", "dependents", "formula_patterns"):
        records.extend(dossier.get(key, []))
    return [
        (str(record.get("node_id") or record.get("id") or "source"), field, formula)
        for record in records
        if isinstance(record, dict)
        for field in ("formula", "r1c1_form")
        if isinstance(formula := record.get(field), str) and formula
    ]


def validate_documentation(text: str, dossier: dict[str, Any]) -> dict[str, Any]:
    """Return source-support evidence and retain the original response verbatim.

    ``qualified`` means an unsupported comparable quotation or processing limit
    needs review. Ellipses and numerical illustrations remain ``not_checkable``
    without an alert. ``unverified`` means no alert was found in this check;
    neither status verifies prose, attribution, topology, values or provenance.
    """
    support: dict[str, set[tuple[str, str]]] = {}
    issues: list[dict[str, Any]] = []
    for source_id, field, formula in _source_formulas(dossier):
        if len(formula) > MAX_FORMULA_CHARS:
            issues.append({"kind": "source_limit", "source_id": source_id})
            continue
        for expression in _expressions(formula):
            support.setdefault(_canonical(expression), set()).add((source_id, field))
    if len(text) > MAX_TEXT_CHARS:
        issues.append({"kind": "response_limit", "characters": len(text)})
    checks = []
    scanned = text[:MAX_TEXT_CHARS]
    wrapper = re.fullmatch(
        r"\s*```(?:markdown|md)\s*\n(.*?)\n```\s*", scanned, re.DOTALL | re.I
    )
    if wrapper:
        # Keep original offsets while exposing inline quotations in a model's
        # optional whole-response Markdown wrapper.
        scanned = " " * wrapper.start(1) + wrapper.group(1)
    for match in _CODE.finditer(scanned):
        quote = match.group(2).strip()
        if len(match.group(1)) == 3:
            quote = re.sub(r"^(?:excel|formula)\s*\n", "", quote, flags=re.I)
        if not (quote.startswith("=") or _FUNCTION.match(quote)):
            continue
        if len(checks) >= MAX_QUOTES:
            issues.append({"kind": "quotation_limit"})
            break
        check: dict[str, Any] = {"quotation": quote, "offset": match.start(2)}
        if len(quote) > MAX_FORMULA_CHARS:
            check["status"] = "not_checkable"
            issues.append({"kind": "formula_length_limit", "check_index": len(checks)})
        elif _incomplete(quote) or re.search(r"\n\s*=", quote):
            check["status"] = "not_checkable"
            check["reason"] = "incomplete_or_illustrative_expression"
        else:
            sources = sorted(support.get(_canonical(quote), []))
            check["status"] = (
                "source_supported"
                if sources
                else "not_source_supported"
                if _has_reference(quote)
                else "not_checkable"
            )
            if check["status"] == "not_checkable":
                check["reason"] = "no_reference_to_compare"
            check["source_ids"] = sorted({source_id for source_id, _ in sources})
            check["evidence"] = [
                {"source_id": source_id, "field": field} for source_id, field in sources
            ]
        checks.append(check)
        if check["status"] == "not_source_supported":
            issues.append({"kind": check["status"], "check_index": len(checks) - 1})
    return {
        "schema_version": 1,
        "status": "qualified" if issues else "unverified",
        "scope": "markdown_formula_quotations_against_sent_formula_and_r1c1_fields",
        "raw_response": text,
        "checks": checks,
        "counts": {
            status: sum(check["status"] == status for check in checks)
            for status in ("source_supported", "not_source_supported", "not_checkable")
        },
        "issues": issues,
        "limitations": [
            (
                "General prose, reference attribution, topology, values and "
                "provenance are not verified."
            ),
            (
                "Source support is lexical, not a proof of semantic correctness "
                "or correct attribution."
            ),
            (
                "Unmatched quotations may be examples, equivalent expressions or "
                "omitted evidence."
            ),
            "Only code quotations beginning with = or a function call are inspected.",
        ],
    }
