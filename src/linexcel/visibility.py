"""Bounded source-visibility adapters for reference-form Excel aggregates.

Contracts: https://support.microsoft.com/en-us/excel/functions/subtotal-function
and https://support.microsoft.com/en-us/excel/functions/aggregate-function.
Only evaluated range arguments reach callbacks. Saved formula values are never
substituted. Filter metadata that cannot distinguish manual hiding is refused
when the requested mode includes manually hidden rows.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass, field

from linexcel.structured_refs import formula_tokens
from linexcel.values import _error_kind, _is_uncomputed

MAX_POSITIONS = 20_000


class VisibilityError(ValueError):
    code = "visibility_unsupported"


@dataclass
class VisibilityPlan:
    formula: str
    functions: list[tuple[str, Callable, int]] = field(default_factory=list)
    evidence: list[dict] = field(default_factory=list)


def _error(kind):
    return {"type": "Error", "kind": kind}


def _aggregate(function, values, ignore_errors, k=None):
    """Reference semantics: text and booleans never become numeric operands."""
    for value in values:
        if _is_uncomputed(value):
            return value
    if ignore_errors:
        values = [value for value in values if _error_kind(value) is None]
    if function == 3:
        return sum(value is not None for value in values)
    if function != 2:
        for value in values:
            if _error_kind(value) is not None:
                return value
    numbers = [
        value
        for value in values
        if isinstance(value, (int, float)) and not isinstance(value, bool)
    ]
    if any(not math.isfinite(value) for value in numbers):
        return _error("Num")
    n = len(numbers)
    if function == 2:
        return n
    if function == 1:
        return math.fsum(numbers) / n if n else _error("Div")
    if function == 4:
        return max(numbers, default=0)
    if function == 5:
        return min(numbers, default=0)
    if function == 6:
        return math.prod(numbers) if n else 0
    if function == 9:
        return math.fsum(numbers)
    if function in {7, 8, 10, 11}:
        sample = function in {7, 10}
        if n < (2 if sample else 1):
            return _error("Div")
        variance = (
            statistics.variance(numbers) if sample else statistics.pvariance(numbers)
        )
        return math.sqrt(variance) if function in {7, 8} else variance
    if function == 12:
        # Empty MEDIAN behavior has no native Excel receipt in this adapter.
        return statistics.median(numbers) if n else _error("NImpl")
    if function == 13:
        counts = Counter(numbers)
        count = max(counts.values(), default=0)
        if count < 2:
            return _error("Na")
        modes = [value for value, occurrences in counts.items() if occurrences == count]
        return modes[0] if len(modes) == 1 else _error("NImpl")
    if _error_kind(k) is not None:
        return k
    if isinstance(k, bool) or not isinstance(k, (int, float)):
        return _error("Value")
    if not math.isfinite(k) or not n:
        return _error("Num")
    ordered = sorted(numbers)
    if function in {14, 15}:
        rank = math.ceil(k)
        if not 1 <= rank <= n:
            return _error("Num")
        return ordered[-rank if function == 14 else rank - 1]
    if function in {17, 19}:
        k = math.trunc(k)
        if not 0 <= k <= 4:
            return _error("Num")
        k /= 4
    exclusive = function in {18, 19}
    if not 0 <= k <= 1 or (exclusive and not 0 < k < 1):
        return _error("Num")
    position = (n + 1) * k - 1 if exclusive else (n - 1) * k
    if not 0 <= position <= n - 1:
        return _error("Num")
    lower = math.floor(position)
    upper = math.ceil(position)
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def _function_name(token):
    return token.value[:-1].upper().removeprefix("_XLFN.")


def _contains(formula, names):
    return any(
        t.type == "FUNC" and t.subtype == "OPEN" and _function_name(t) in names
        for t in formula_tokens(formula or "=0")
    )


def _literal_integer(tokens):
    meaningful = [t for t in tokens if t.type != "WHITE-SPACE"]
    if len(meaningful) != 1 or meaningful[0].subtype != "NUMBER":
        raise VisibilityError("Aggregate function and option must be literal integers.")
    try:
        value = float(meaningful[0].value)
    except ValueError as exc:
        raise VisibilityError("Invalid aggregate option.") from exc
    if not math.isfinite(value) or not value.is_integer():
        raise VisibilityError("Aggregate function and option must be literal integers.")
    return int(value)


def _callback(function, masks, ignore_errors, has_k):
    def calculate(*arguments):
        values = []
        for argument, (height, width, excluded) in zip(arguments, masks):
            matrix = argument if isinstance(argument, list) else [[argument]]
            if len(matrix) != height or any(
                not isinstance(row, list) or len(row) != width for row in matrix
            ):
                return _error("Value")
            values.extend(
                value
                for row, items in enumerate(matrix)
                for col, value in enumerate(items)
                if row * width + col not in excluded
            )
        try:
            result = _aggregate(
                function, values, ignore_errors, arguments[-1] if has_k else None
            )
            if isinstance(result, (int, float)) and not math.isfinite(result):
                return _error("Num")
            return result
        except (OverflowError, ValueError, statistics.StatisticsError):
            return _error("Num")

    return calculate


def rewrite_visibility_formula(
    formula: str, sheet: str, resolve_reference
) -> VisibilityPlan:
    """Resolver supplies expression, bounds, cells and authoritative visibility.

    ``cells`` is a sparse iterable of source-node dictionaries, with row,
    column, formula and context.rowHidden. ``activeFilters`` is conservative;
    ``visibilityKnown`` must be true. Reference expressions must be fully
    resolved before registration; unsupported resolver inputs raise ValueError.
    """
    tokens = formula_tokens(formula)
    plan = VisibilityPlan(formula)
    replacements = []
    stack = []
    for index, token in enumerate(tokens):
        if token.subtype == "OPEN" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            stack.append((index, []))
        elif token.type == "SEP" and token.subtype == "ARG" and stack:
            stack[-1][1].append(index)
        elif token.subtype == "CLOSE" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            if not stack:
                raise VisibilityError("Unbalanced aggregate formula.")
            start, separators = stack.pop()
            opening = tokens[start]
            if opening.type != "FUNC":
                continue
            name = _function_name(opening)
            if name not in {"SUBTOTAL", "AGGREGATE"}:
                continue
            boundaries = [start, *separators, index]
            args = [tokens[a + 1 : b] for a, b in zip(boundaries, boundaries[1:])]
            required = 2 if name == "SUBTOTAL" else 3
            if len(args) < required:
                raise VisibilityError(f"{name} requires reference arguments.")
            number = _literal_integer(args[0])
            if name == "SUBTOTAL":
                if number not in {*range(1, 12), *range(101, 112)}:
                    raise VisibilityError("SUBTOTAL function number is unsupported.")
                hidden = number >= 100
                function, errors, nested = number % 100, False, {"SUBTOTAL"}
                reference_args = args[1:]
            else:
                option = _literal_integer(args[1])
                if not 1 <= number <= 19 or not 0 <= option <= 7:
                    raise VisibilityError("AGGREGATE function/option is unsupported.")
                hidden, errors = option in {1, 3, 5, 7}, option in {2, 3, 6, 7}
                nested = {"SUBTOTAL", "AGGREGATE"} if option < 4 else set()
                function, reference_args = number, args[2:]
            has_k = function >= 14
            k_tokens = reference_args[-1] if has_k else None
            if has_k:
                if len(reference_args) != 2:
                    raise VisibilityError(
                        "This AGGREGATE form requires one range and k."
                    )
                reference_args = reference_args[:-1]
                assert k_tokens is not None
                clean_k = [t for t in k_tokens if t.type != "WHITE-SPACE"]
                if len(clean_k) != 1 or clean_k[0].subtype != "NUMBER":
                    raise VisibilityError("AGGREGATE k must be a numeric literal.")
                k_value = float(clean_k[0].value)
                if not math.isfinite(k_value) or (
                    function in {14, 15, 17, 19} and not k_value.is_integer()
                ):
                    raise VisibilityError("Fractional ranks/quartiles are unsupported.")
            if len(reference_args) > 253:
                raise VisibilityError("Aggregate reference count exceeds 253.")
            expressions, masks, evidence = [], [], []
            positions = 0
            for argument in reference_args:
                clean = [t for t in argument if t.type != "WHITE-SPACE"]
                if (
                    len(clean) != 1
                    or clean[0].type != "OPERAND"
                    or clean[0].subtype != "RANGE"
                ):
                    raise VisibilityError(
                        "Only resolved reference-form aggregates are supported."
                    )
                metadata = resolve_reference(clean[0].value, sheet)
                if not metadata.get("visibilityKnown"):
                    raise VisibilityError("Source row visibility is unavailable.")
                if metadata.get("activeFilters") and not hidden:
                    raise VisibilityError(
                        "Active filters cannot distinguish manually hidden rows "
                        "in this mode."
                    )
                r1, c1, r2, c2 = metadata["bounds"]
                height, width = r2 - r1 + 1, c2 - c1 + 1
                positions += height * width
                if height < 1 or width < 1 or positions > MAX_POSITIONS:
                    raise VisibilityError(
                        "Aggregate visibility exceeds 20000 positions."
                    )
                excluded = {
                    (row - r1) * width + column
                    for row in metadata.get("hiddenRows", ())
                    if hidden and r1 <= row <= r2
                    for column in range(width)
                }
                for count, cell in enumerate(metadata["cells"], 1):
                    if count > MAX_POSITIONS:
                        raise VisibilityError(
                            "Aggregate source metadata exceeds its limit."
                        )
                    row, col = cell["row"], cell["column"]
                    if not (r1 <= row <= r2 and c1 <= col <= c2):
                        raise VisibilityError(
                            "Aggregate metadata is outside its reference."
                        )
                    if (hidden and cell.get("context", {}).get("rowHidden")) or (
                        nested and _contains(cell.get("formula"), nested)
                    ):
                        excluded.add((row - r1) * width + col - c1)
                expressions.append(metadata["expression"])
                masks.append((height, width, excluded))
                evidence.append(
                    {
                        "reference": clean[0].value,
                        "expression": metadata["expression"],
                        "bounds": list(metadata["bounds"]),
                        "excludedPositions": sorted(excluded),
                        "activeFilters": bool(metadata.get("activeFilters")),
                    }
                )
            if has_k:
                if not k_tokens:
                    raise VisibilityError("AGGREGATE k argument is missing.")
                expressions.append("".join(t.value for t in k_tokens))
            identity = {
                "sheet": sheet,
                "formula": formula,
                "start": start,
                "ranges": evidence,
            }
            digest = hashlib.sha256(
                json.dumps(identity, sort_keys=True).encode()
            ).hexdigest()[:24]
            custom = "LINEXCEL.VISIBILITY.V" + digest
            replacement = custom + "(" + ",".join(expressions) + ")"
            if any(a > start and b < index for a, b, _ in replacements):
                raise VisibilityError(
                    "Nested aggregate expressions require direct source references."
                )
            replacements.append((start, index, replacement))
            plan.functions.append(
                (custom, _callback(function, masks, errors, has_k), len(expressions))
            )
            plan.evidence.append(
                {"function": name, "number": number, "ranges": evidence}
            )
    if stack:
        raise VisibilityError("Unbalanced aggregate formula.")
    for start, end, replacement in sorted(replacements, reverse=True):
        tokens[start].value = replacement
        del tokens[start + 1 : end + 1]
    plan.formula = "=" + "".join(t.value for t in tokens)
    return plan


def register_visibility_functions(
    workbook, plan: VisibilityPlan, registered=None, nonresults=None
):
    for name, callback, count in plan.functions:
        if registered is not None and name in registered:
            continue

        def observed(*arguments, _callback=callback, _name=name):
            value = _callback(*arguments)
            if nonresults is not None and _is_uncomputed(value):
                nonresults.add(_name)
            return value

        workbook.register_function(name, observed, min_args=count, max_args=count)
        if registered is not None:
            registered.add(name)
