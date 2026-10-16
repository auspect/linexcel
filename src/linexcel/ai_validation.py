"""Bounded quotation checks, never a factual certification of generated prose.

Only Markdown code quotations of complete Excel formulas/function expressions
are compared with the exact source dossier sent to the model. A missing match
is a qualification, not proof of an error: examples, equivalent expressions,
and omitted source evidence can all legitimately fail this lexical check.
"""

from __future__ import annotations

import hashlib
import json
import re
from decimal import Context, Decimal, DecimalException, localcontext
from fractions import Fraction
from operator import add, eq, ge, gt, le, lt, mul, ne, sub, truediv
from typing import Any

MAX_TEXT_CHARS = 200_000
MAX_FORMULA_CHARS = 8192
MAX_QUOTES = 256
MAX_CLAIMS = 96
MAX_CLAIM_BLOCK_CHARS = 32_000
REVIEW_VERSION = "linexcel-factual-review-v5"
_CLAIMS = re.compile(r"```json_claims\s*(.*?)```", re.DOTALL | re.IGNORECASE)
_CLAIM_START = re.compile(r"```json_claims\b", re.IGNORECASE)
_TABLES = re.compile(r"```json_tables\s*(.*?)```", re.DOTALL | re.IGNORECASE)
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
    for key in (
        "precedents",
        "dependents",
        "formula_patterns",
        "sample",
        "value_samples",
    ):
        items = dossier.get(key)
        if isinstance(items, list):
            records.extend(items)
    if isinstance(dossier.get("node"), dict):
        records.append(dossier["node"])
    return [
        (
            str(
                record.get("node_id")
                or record.get("id")
                or record.get("addr")
                or "source"
            ),
            field,
            formula,
        )
        for record in records
        if isinstance(record, dict)
        for field in (
            "formula",
            "source_formula",
            "canonical_formula",
            "sourceFormula",
            "canonicalFormula",
            "r1c1_form",
        )
        if isinstance(formula := record.get(field), str) and formula
    ]


def validate_documentation(
    text: str,
    dossier: dict[str, Any],
    *,
    quotation_dossier: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Return source-support evidence and retain the original response verbatim.

    ``qualified`` means an unsupported comparable quotation or processing limit
    needs review. Ellipses and numerical illustrations remain ``not_checkable``
    without an alert. ``unverified`` means no alert was found in this check;
    neither status verifies prose, attribution, topology, values or provenance.
    """
    support: dict[str, set[tuple[str, str]]] = {}
    issues: list[dict[str, Any]] = []
    for source_id, field, formula in _source_formulas(
        quotation_dossier if quotation_dossier is not None else dossier
    ):
        if len(formula) > MAX_FORMULA_CHARS:
            issues.append({"kind": "source_limit", "source_id": source_id})
            continue
        for expression in _expressions(formula):
            support.setdefault(_canonical(expression), set()).add((source_id, field))
    if len(text) > MAX_TEXT_CHARS:
        issues.append({"kind": "response_limit", "characters": len(text)})
    checks = []
    # Machine evidence bindings are inspected separately, not as formula prose.
    scanned = _CLAIMS.sub(lambda match: " " * len(match.group()), text[:MAX_TEXT_CHARS])
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
    claim_report = validate_claims(text, dossier)
    issues.extend(claim_report["issues"])
    return {
        "schema_version": 2,
        "status": "qualified" if issues else "unverified",
        "scope": (
            "markdown_formula_quotations_against_sent_original_canonical_and_r1c1_fields"
        ),
        "formula_field_roles": {
            "source_formula": "presented_original_or_derived_shared_translation",
            "sourceFormula": "presented_original_or_derived_shared_translation",
            "formula_provenance.xmlText": "decoded_stored_xml_without_added_equals",
            "formulaProvenance.xmlText": "decoded_stored_xml_without_added_equals",
            "canonical_formula": "canonical_expression",
            "canonicalFormula": "canonical_expression",
            "formula": "legacy_expression_not_verbatim_xml_proof",
            "r1c1_form": "derived_relative_expression",
        },
        "raw_response": text,
        "checks": checks,
        "counts": {
            status: sum(check["status"] == status for check in checks)
            for status in ("source_supported", "not_source_supported", "not_checkable")
        },
        "issues": issues,
        "factual_review": claim_report,
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


def review_protocol() -> dict[str, Any]:
    """Version the evidence contract without duplicating the dossier's values."""
    return {
        "version": REVIEW_VERSION,
        "claims_required": True,
        "claims": "json_claims: statement, classification, evidence[{pointer,value}]",
        "calculation": (
            "optional {operator, operands:[{pointer,value},{pointer,value}], result}; "
            "numeric JSON operands only, bounded decimal arithmetic/comparison, "
            "not Excel verification"
        ),
        "pointer_basis": "RFC6901 pointers into this exact sent JSON dossier",
        "scope": (
            "Typed evidence bindings only; prose meaning and pixels need "
            "independent review"
        ),
    }


def strip_claims(text: str) -> str:
    """Remove private review scaffolding from visible prose; retain it in reports."""
    prose = _CLAIMS.sub("", text)
    incomplete = _CLAIM_START.search(prose)
    return (prose[: incomplete.start()] if incomplete else prose).strip()


def _pointer(dossier: Any, pointer: str) -> tuple[bool, Any]:
    if (
        not isinstance(pointer, str)
        or not pointer.startswith("/")
        or len(pointer) > 512
    ):
        return False, None
    item = dossier
    for raw in pointer[1:].split("/"):
        if re.search(r"~(?![01])", raw):
            return False, None
        key = raw.replace("~1", "/").replace("~0", "~")
        if isinstance(item, dict):
            # An omission marker is unknown, not a factual source to quote.
            if item.get("omitted") or key not in item or key == "review_protocol":
                return False, None
            item = item[key]
        elif isinstance(item, list) and re.fullmatch(r"0|[1-9][0-9]*", key):
            index = int(key)
            if index >= len(item):
                return False, None
            item = item[index]
        else:
            return False, None
    if isinstance(item, dict) and item.get("omitted"):
        return False, None
    return True, item


def _typed_equal(left: Any, right: Any) -> bool:
    # JSON numbers share a type; booleans must never compare equal to 0/1.
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left == right
    if type(left) is not type(right):
        return False
    if isinstance(left, dict):
        return left.keys() == right.keys() and all(
            _typed_equal(left[k], right[k]) for k in left
        )
    if isinstance(left, list):
        return len(left) == len(right) and all(
            _typed_equal(a, b) for a, b in zip(left, right)
        )
    return left == right


def _calculation_check(calculation: Any, dossier: dict[str, Any]) -> dict[str, Any]:
    """Check explicitly claimed scalar arithmetic, never infer a prose meaning.

    Operands must match numeric values in the exact sent dossier. This bounded
    Decimal check is neither Excel coercion/rounding nor an independent oracle.
    """
    operations = {
        "+": add,
        "-": sub,
        "*": mul,
        "/": truediv,
        ">": gt,
        ">=": ge,
        "<": lt,
        "<=": le,
        "=": eq,
        "<>": ne,
    }
    result = {
        "scope": "explicit numeric claim over bound JSON operands; not Excel evaluation"
    }
    if not isinstance(calculation, dict):
        return {
            **result,
            "verdict": "indeterminate",
            "reason": "invalid_calculation_schema",
        }
    op, operands = calculation.get("operator"), calculation.get("operands")
    if (
        not isinstance(op, str)
        or op not in operations
        or not isinstance(operands, list)
        or len(operands) != 2
    ):
        return {
            **result,
            "verdict": "indeterminate",
            "reason": "invalid_calculation_schema",
        }
    numbers = []
    for operand in operands:
        if not isinstance(operand, dict):
            return {
                **result,
                "verdict": "indeterminate",
                "reason": "invalid_calculation_operand",
            }
        value = operand.get("value")
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or len(str(value)) > 64
        ):
            return {
                **result,
                "verdict": "indeterminate",
                "reason": "non_numeric_calculation_operand",
            }
        found, actual = _pointer(dossier, operand.get("pointer"))
        if not found:
            return {
                **result,
                "verdict": "indeterminate",
                "reason": "missing_calculation_operand",
            }
        if not _typed_equal(value, actual):
            return {
                **result,
                "verdict": "contradicted",
                "reason": "calculation_operand_mismatch",
                "pointer": operand.get("pointer"),
                "evidence_value": actual,
            }
        number = Decimal(str(value))
        if not number.is_finite():
            return {
                **result,
                "verdict": "indeterminate",
                "reason": "non_finite_calculation_operand",
            }
        numbers.append(number)
    try:
        if op == "/" and numbers[1]:
            denominator = (Fraction(numbers[0]) / Fraction(numbers[1])).denominator
            for factor in (2, 5):
                while denominator % factor == 0:
                    denominator //= factor
            if denominator != 1:
                return {
                    **result,
                    "verdict": "indeterminate",
                    "reason": "nonterminating_decimal_division",
                }
        # Independent of the embedding application's ambient Decimal context.
        with localcontext(Context(prec=1000)):
            expected = operations[op](*numbers)
        claimed = calculation.get("result")
        if isinstance(expected, bool):
            matches = type(claimed) is bool and claimed == expected
        else:
            matches = (
                not isinstance(claimed, bool)
                and isinstance(claimed, (int, float))
                and Decimal(str(claimed)) == expected
            )
    except (DecimalException, ValueError, TypeError, OverflowError):
        return {
            **result,
            "verdict": "indeterminate",
            "reason": "undefined_or_unbounded_calculation",
        }
    return {
        **result,
        "verdict": "supported" if matches else "contradicted",
        "expected_result": expected if isinstance(expected, bool) else str(expected),
        "claimed_result": claimed,
    }


def validate_claims(text: str, dossier: dict[str, Any]) -> dict[str, Any]:
    """Check explicit atomic claims against the exact sent evidence, never meaning.

    The generator is not an independent oracle. A supported binding proves only
    that its typed value was present at the stated pointer. Human/independent
    review must still check attribution, the statement's meaning, completeness,
    original workbook XML and visible image details.
    """
    blocks = list(_CLAIMS.finditer(text[:MAX_TEXT_CHARS]))
    checks: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    prose = _TABLES.sub("", strip_claims(text[:MAX_TEXT_CHARS]))
    if _CLAIM_START.search(_CLAIMS.sub("", text[:MAX_TEXT_CHARS])):
        issues.append({"kind": "incomplete_claim_block"})
    contract = "present" if blocks else "missing"
    required = (
        dossier.get("review_protocol", {}).get("claims_required") is True
        if isinstance(dossier.get("review_protocol"), dict)
        else False
    )
    if required and not blocks:
        issues.append({"kind": "missing_required_claim_block"})
    if required and len(blocks) > 1:
        issues.append({"kind": "multiple_required_claim_blocks"})
    claims = []
    for block in blocks:
        if len(block[1]) > MAX_CLAIM_BLOCK_CHARS:
            issues.append({"kind": "claim_block_limit"})
            continue
        try:
            payload = json.loads(block[1])
        except (ValueError, RecursionError):
            issues.append({"kind": "invalid_claim_json"})
            continue
        if not isinstance(payload, list):
            issues.append({"kind": "invalid_claim_schema"})
            continue
        available = MAX_CLAIMS - len(claims)
        claims.extend(payload[:available])
        if len(payload) > available:
            issues.append({"kind": "claim_count_limit"})
            break
    if required and blocks and not claims:
        issues.append({"kind": "empty_required_claim_ledger"})
    for claim in claims:
        if not isinstance(claim, dict):
            issues.append({"kind": "invalid_claim_schema"})
            continue
        statement = claim.get("statement")
        classification = claim.get("classification")
        evidence = claim.get("evidence")
        check = {
            "statement": statement,
            "classification": classification,
            "bindings": [],
        }
        if (
            not isinstance(statement, str)
            or not statement
            or len(statement) > 2000
            or not isinstance(classification, str)
            or classification not in {"fact", "inference", "unknown"}
            or not isinstance(evidence, list)
            or len(evidence) > 8
        ):
            check.update(verdict="indeterminate", reason="invalid_claim_schema")
            issues.append({"kind": "invalid_claim_schema", "claim_index": len(checks)})
        elif statement not in prose:
            check.update(verdict="indeterminate", reason="statement_not_in_response")
            issues.append(
                {"kind": "statement_not_in_response", "claim_index": len(checks)}
            )
        else:
            for binding in evidence:
                if not isinstance(binding, dict) or "value" not in binding:
                    check["bindings"].append(
                        {"verdict": "indeterminate", "reason": "invalid_binding"}
                    )
                    continue
                found, value = _pointer(dossier, binding.get("pointer"))
                verdict = (
                    "indeterminate"
                    if not found
                    else "supported"
                    if _typed_equal(value, binding["value"])
                    else "contradicted"
                )
                check["bindings"].append(
                    {
                        "pointer": binding.get("pointer"),
                        "claimed_value": binding["value"],
                        "verdict": verdict,
                        **(
                            {"evidence_value": value}
                            if found
                            else {"reason": "missing_or_omitted_evidence"}
                        ),
                    }
                )
            verdicts = [binding["verdict"] for binding in check["bindings"]]
            check["verdict"] = (
                "contradicted"
                if "contradicted" in verdicts
                else "supported"
                if classification == "fact"
                and verdicts
                and all(v == "supported" for v in verdicts)
                else "indeterminate"
            )
            if check["verdict"] == "contradicted" or (
                classification == "fact" and check["verdict"] != "supported"
            ):
                issues.append(
                    {"kind": "claim_evidence_mismatch", "claim_index": len(checks)}
                )
            if "calculation" in claim:
                calculation = _calculation_check(claim["calculation"], dossier)
                check["calculation"] = calculation
                if calculation["verdict"] != "supported":
                    issues.append(
                        {
                            "kind": "calculation_claim_mismatch",
                            "claim_index": len(checks),
                            "reason": calculation.get("reason", "result_mismatch"),
                        }
                    )
                    check["verdict"] = calculation["verdict"]
        checks.append(check)
    return {
        "version": REVIEW_VERSION,
        "contract": contract,
        "status": "needs_review"
        if issues
        else "bindings_checked"
        if checks
        else "not_supplied",
        "dossier_sha256": hashlib.sha256(
            json.dumps(
                dossier, ensure_ascii=False, sort_keys=True, default=str
            ).encode()
        ).hexdigest(),
        "checks": checks,
        "issues": issues,
        "counts": {
            verdict: sum(c["verdict"] == verdict for c in checks)
            for verdict in ("supported", "contradicted", "indeterminate")
        },
        "limitations": [
            "Bindings do not certify the associated prose, "
            "its attribution or exhaustiveness.",
            "Missing claims, inferences and unknowns remain unverified; "
            "images need a separate pixel review.",
        ],
    }
