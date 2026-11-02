"""Bounded source-derived INDIRECT/OFFSET resolution without saved formula values.

Address expressions may use source literals and arithmetic/concatenation formulas.
Unsupported functions, cycles, ambiguous coercions and external dynamic contexts
are refused. The resulting native formula and dependencies are separately recorded.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass

from linexcel.refs import Rect, parse_ref, split_sheet_prefix
from linexcel.structured_refs import formula_tokens


@dataclass
class DynamicPlan:
    formula: str
    evidence: list[dict]


def _number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("Dynamic address arithmetic requires numeric source values.")
    if not math.isfinite(value):
        raise ValueError("Dynamic address arithmetic produced a non-finite value.")
    return value


def _text(value):
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int) or isinstance(value, float) and value.is_integer():
        return str(int(value))
    raise ValueError("Dynamic address concatenation requires text or integral values.")


def scalar_expression(expression, sheet, resolve, active=()):
    """Evaluate a deliberately small, checked Excel scalar grammar."""
    if len(expression) > 8192 or len(active) > 32:
        raise ValueError("Dynamic address expression exceeds its resolution budget.")
    tokens = [t for t in formula_tokens("=" + expression) if t.type != "WHITE-SPACE"]
    if len(tokens) > 1024:
        raise ValueError("Dynamic address expression exceeds its token budget.")
    position = 0

    def primary():
        nonlocal position
        if position >= len(tokens):
            raise ValueError("Incomplete dynamic address expression.")
        token = tokens[position]
        position += 1
        if token.type == "OPERATOR-PREFIX" and token.value in {"+", "-"}:
            value = _number(primary())
            return value if token.value == "+" else -value
        if token.type == "PAREN" and token.subtype == "OPEN":
            value = expression_value(0)
            if position >= len(tokens) or tokens[position].subtype != "CLOSE":
                raise ValueError("Unbalanced dynamic address expression.")
            position += 1
            return value
        if token.type != "OPERAND":
            raise ValueError(
                "Unsupported function/operator in dynamic address expression."
            )
        if token.subtype == "TEXT":
            return token.value[1:-1].replace('""', '"')
        if token.subtype == "NUMBER":
            return _number(float(token.value))
        if token.subtype == "LOGICAL":
            return token.value.upper() == "TRUE"
        if token.subtype == "RANGE":
            node, _ = resolve(token.value, sheet)
            if node["id"] in active or node.get("diagnostics"):
                raise ValueError("Dynamic address expression is cyclic or unsupported.")
            if node.get("kind") == "input":
                value = node.get("cached_comparison_value", node.get("value"))
                if node.get("sourceDataType") == "e" or value is None:
                    raise ValueError(
                        "Dynamic address expression contains an error or blank."
                    )
                return value
            if node.get("kind") == "cell" or node.get("kind") == "name":
                return scalar_expression(
                    node["formula"][1:],
                    node.get("sheet") or sheet,
                    resolve,
                    (*active, node["id"]),
                )
        raise ValueError("Dynamic address expression is not a supported scalar.")

    precedence = {"&": 1, "+": 2, "-": 2, "*": 3, "/": 3}

    def expression_value(minimum):
        nonlocal position
        left = primary()
        while position < len(tokens):
            operator = tokens[position].value
            priority = precedence.get(operator, -1)
            if priority < minimum:
                break
            position += 1
            right = expression_value(priority + 1)
            if operator == "&":
                left = _text(left) + _text(right)
            else:
                a, b = _number(left), _number(right)
                if operator == "+":
                    left = a + b
                elif operator == "-":
                    left = a - b
                elif operator == "*":
                    left = a * b
                else:
                    if b == 0:
                        raise ValueError(
                            "Division by zero in dynamic address expression."
                        )
                    left = a / b
                _number(left)
        return left

    result = expression_value(0)
    if position != len(tokens):
        raise ValueError("Unsupported syntax in dynamic address expression.")
    return result


def _r1c1(text, sheet, row, column):
    prefix, body = split_sheet_prefix(text)
    parts = body.upper().split(":")
    if len(parts) > 2:
        raise ValueError("Invalid R1C1 dynamic reference.")
    coordinates = []
    for part in parts:
        match = re.fullmatch(r"R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)", part)
        if not match:
            raise ValueError("Unsupported R1C1 dynamic reference.")
        coordinates.append(
            tuple(
                base + int(value[1:-1])
                if value.startswith("[")
                else int(value)
                if value
                else base
                for value, base in zip(match.groups(), (row, column), strict=True)
            )
        )
    first, last = coordinates[0], coordinates[-1]
    return Rect(prefix or sheet, first[0], first[1], last[0], last[1])


def rewrite_dynamic_formula(formula, sheet, row, column, resolve):
    tokens = formula_tokens(formula)
    evidence = []
    for _ in range(64):
        stack = []
        found = None
        for index, token in enumerate(tokens):
            if token.type == "FUNC" and token.subtype == "OPEN":
                stack.append(index)
            elif token.type == "FUNC" and token.subtype == "CLOSE":
                start = stack.pop()
                function = tokens[start].value[:-1].upper().removeprefix("_XLFN.")
                if function in {"INDIRECT", "OFFSET"}:
                    found = (start, index, function)
                    break
        if found is None:
            return DynamicPlan("=" + "".join(t.value for t in tokens), evidence)
        start, end, function = found
        arguments = [[]]
        depth = 0
        for token in tokens[start + 1 : end]:
            if token.subtype == "OPEN":
                depth += 1
            elif token.subtype == "CLOSE":
                depth -= 1
            if token.type == "SEP" and token.subtype == "ARG" and depth == 0:
                arguments.append([])
            else:
                arguments[-1].append(token.value)
        arguments = ["".join(argument).strip() for argument in arguments]

        def scalar(text):
            return scalar_expression(text, sheet, resolve)

        if function == "INDIRECT":
            if len(arguments) not in {1, 2}:
                raise ValueError("INDIRECT requires one or two arguments.")
            address = scalar(arguments[0])
            style = scalar(arguments[1]) if len(arguments) == 2 else True
            if not isinstance(address, str) or style not in {True, False, 0, 1}:
                raise ValueError(
                    "INDIRECT address/style must be unambiguous source values."
                )
            if style:
                _, expression = resolve(address, sheet)
                rectangle = parse_ref(expression, default_sheet=sheet)
            else:
                rectangle = _r1c1(address, sheet, row, column)
        else:
            if not 3 <= len(arguments) <= 5:
                raise ValueError("OFFSET requires three to five arguments.")
            _, expression = resolve(arguments[0], sheet)
            base = parse_ref(expression, default_sheet=sheet)
            if base is None:
                raise ValueError("OFFSET base must be a contiguous source reference.")
            values = [_number(scalar(text)) if text else None for text in arguments[1:]]
            if any(
                value is not None and not float(value).is_integer() for value in values
            ):
                raise ValueError("OFFSET requires integral offsets and dimensions.")
            dr, dc = values[:2]
            if dr is None or dc is None:
                raise ValueError("OFFSET requires both offsets.")
            height = (
                values[2]
                if len(values) > 2 and values[2] is not None
                else base.r2 - base.r1 + 1
            )
            width = (
                values[3]
                if len(values) > 3 and values[3] is not None
                else base.c2 - base.c1 + 1
            )
            if height <= 0 or width <= 0:
                raise ValueError("OFFSET dimensions must be positive.")
            rectangle = Rect(
                base.sheet,
                base.r1 + int(dr),
                base.c1 + int(dc),
                base.r1 + int(dr + height) - 1,
                base.c1 + int(dc + width) - 1,
            )
        if rectangle is None or not (
            1 <= rectangle.r1 <= rectangle.r2 <= 1048576
            and 1 <= rectangle.c1 <= rectangle.c2 <= 16384
        ):
            raise ValueError("Dynamic reference exceeds Excel worksheet bounds.")
        _, expression = resolve(rectangle.to_a1(), sheet)
        original = "".join(t.value for t in tokens[start : end + 1])
        replacement = formula_tokens("=" + expression)
        tokens[start : end + 1] = replacement
        evidence.append(
            {
                "function": function,
                "sourceExpression": original,
                "resolvedReference": expression,
                "basis": "source_literals_and_checked_scalar_expressions",
            }
        )
    raise ValueError("Dynamic reference expansion exceeds its call budget.")
