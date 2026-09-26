"""Typed Excel semantics at the native engine boundary.

These callbacks consume evaluated arguments, so corrected results participate
in the native dependency graph. They never read saved formula caches or reenter
the workbook. Formula rewriting changes evaluation only, not source evidence.
"""

from __future__ import annotations

import datetime as dt
import io
import math
import re
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

import formualizer as fz

from linexcel.refs import parse_ref, split_sheet_prefix
from linexcel.values import _error_kind

_CONCAT = "LINEXCEL.COMPAT.CONCAT"
_COMPARE = "LINEXCEL.COMPAT.COMPARE"
_COMPARISONS = frozenset({"=", "<>", "<", ">", "<=", ">="})


class NameBindings(dict):
    """Definitions with the valid source sheet namespace retained."""

    def __init__(self, sheet_names=()):
        super().__init__()
        self.sheet_names = {name.casefold(): name for name in sheet_names}


def _absolute_reference(reference: str) -> bool:
    _, body = split_sheet_prefix(reference)
    return bool(
        re.fullmatch(
            r"\$[A-Za-z]{1,3}\$[0-9]+(?::\$[A-Za-z]{1,3}\$[0-9]+)?"
            r"|\$[A-Za-z]{1,3}:\$[A-Za-z]{1,3}|\$[0-9]+:\$[0-9]+",
            body,
        )
    )


def _error(kind: str) -> dict[str, str]:
    return {"type": "Error", "kind": kind}


def _matrix(value: Any) -> list[list[Any]]:
    return value if isinstance(value, list) else [[value]]


def _shape(rows: list[list[Any]]) -> tuple[int, int] | None:
    if not rows or not isinstance(rows[0], list) or not rows[0]:
        return None
    width = len(rows[0])
    if any(not isinstance(row, list) or len(row) != width for row in rows):
        return None
    return len(rows), width


def _sumproduct(*arguments: Any) -> Any:
    """Separate arrays ignore text/bools; operators inside arguments already ran.

    Microsoft: https://support.microsoft.com/en-us/excel/functions/sumproduct-function
    Array dimensions must match; nonnumeric entries count as zero. Errors still
    propagate even when a corresponding multiplicand is zero or nonnumeric.
    """
    matrices = [_matrix(arg) for arg in arguments]
    dimensions = [_shape(matrix) for matrix in matrices]
    if (
        not dimensions
        or dimensions[0] is None
        or any(shape != dimensions[0] for shape in dimensions)
    ):
        return _error("Value")
    height, width = dimensions[0]
    products = []
    for row in range(height):
        for column in range(width):
            product = 1.0
            for matrix in matrices:
                value = matrix[row][column]
                if _error_kind(value) is not None:
                    return value
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    value = 0
                product *= value
            products.append(product)
    try:
        total = math.fsum(products)
    except (OverflowError, ValueError):
        return _error("Num")
    return total if math.isfinite(total) else _error("Num")


def _date_serial(value: Any, date_system: str) -> Any:
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        from openpyxl.utils.datetime import (
            CALENDAR_MAC_1904,
            CALENDAR_WINDOWS_1900,
            to_excel,
        )

        epoch = CALENDAR_MAC_1904 if date_system == "1904" else CALENDAR_WINDOWS_1900
        return to_excel(value, epoch)
    return value


def _excel_date(year: Any, month: Any, day: Any, date_system: str) -> Any:
    """DATE serial arithmetic, including Excel's intentionally fictitious day 60.

    Year/month/day normalization follows Microsoft's DATE specification; the
    lower bound and the 1900 leap-day cases are also checked against Excel.
    Integer calendar arithmetic permits month/day overflow before bounds apply.
    """
    parts = []
    for value in (year, month, day):
        if _error_kind(value) is not None:
            return value
        try:
            number = 0.0 if value is None else float(value)
        except (ValueError, TypeError):
            return _error("Value")
        if not math.isfinite(number):
            return _error("Num")
        parts.append(math.trunc(number))
    year, month, day = parts
    if not 0 <= year <= 9999:
        return _error("Num")
    if year < 1900:
        year += 1900
    offset, month = divmod(month - 1, 12)
    year += offset
    month += 1
    previous = year - 1
    ordinal = 365 * previous + previous // 4 - previous // 100 + previous // 400
    ordinal += (367 * month - 362) // 12
    if month > 2:
        leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        ordinal -= 1 if leap else 2
    serial = ordinal - dt.date(1899, 12, 31).toordinal() + day
    if (year, month) >= (1900, 3):
        serial += 1
    if date_system == "1904":
        serial -= 1462
    maximum = 2957003 if date_system == "1904" else 2958465
    return serial if 0 <= serial <= maximum else _error("Num")


def _date_array(year: Any, month: Any, day: Any, date_system: str) -> Any:
    if not any(isinstance(value, list) for value in (year, month, day)):
        return _excel_date(year, month, day, date_system)
    matrices = [_matrix(value) for value in (year, month, day)]
    shapes = [_shape(matrix) for matrix in matrices]
    if any(shape is None for shape in shapes):
        return _error("Value")
    shapes = [shape for shape in shapes if shape is not None]
    height, width = (max(shape[i] for shape in shapes) for i in (0, 1))
    if any(h not in (1, height) or w not in (1, width) for h, w in shapes):
        return _error("Value")

    def cell(row, column):
        values = [
            matrix[row % h][column % w] for matrix, (h, w) in zip(matrices, shapes)
        ]
        return _excel_date(values[0], values[1], values[2], date_system)

    return [[cell(r, c) for c in range(width)] for r in range(height)]


def _date_part(value: Any, part: str, date_system: str) -> Any:
    if isinstance(value, list):
        return [[_date_part(item, part, date_system) for item in row] for row in value]
    if _error_kind(value) is not None:
        return value
    value = _date_serial(value, date_system)
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        try:
            value = _date_serial(dt.date.fromisoformat(value), date_system)
        except ValueError:
            pass
    if isinstance(value, str):
        try:
            value = float(value)
        except ValueError:
            pass
    if isinstance(value, str):
        # Keep native date-text parsing, whose locale policy belongs to the
        # general engine. This separate workbook cannot reenter our callback.
        config = fz.EvaluationConfig()
        config.date_system = date_system
        fallback = fz.Workbook(config=fz.WorkbookConfig(eval_config=config))
        fallback.add_sheet("S")
        fallback.set_value("S", 1, 1, value)
        fallback.set_formula("S", 1, 2, f"={part}(A1)")
        return fallback.evaluate_cell("S", 1, 2)
    if value is None:
        value = 0
    if not isinstance(value, (int, float)) or not math.isfinite(value):
        return _error("Value")
    if value < 0 or value >= (2957004 if date_system == "1904" else 2958466):
        return _error("Num")
    day = math.floor(value)
    if date_system == "1900" and day in (0, 60):
        return {
            "DAY": 0 if day == 0 else 29,
            "MONTH": 1 if day == 0 else 2,
            "YEAR": 1900,
        }[part]
    epoch = dt.date(1904, 1, 1) if date_system == "1904" else dt.date(1899, 12, 31)
    date = epoch + dt.timedelta(days=day - (date_system == "1900" and day > 60))
    return {"DAY": date.day, "MONTH": date.month, "YEAR": date.year}[part]


def _binary_array(function, left: Any, right: Any) -> Any:
    if not isinstance(left, list) and not isinstance(right, list):
        return function(left, right)
    a, b = _matrix(left), _matrix(right)
    sa, sb = _shape(a), _shape(b)
    if (
        sa is None
        or sb is None
        or any(x != y and x != 1 and y != 1 for x, y in zip(sa, sb))
    ):
        return _error("Value")
    return [
        [
            function(a[r % sa[0]][c % sa[1]], b[r % sb[0]][c % sb[1]])
            for c in range(max(sa[1], sb[1]))
        ]
        for r in range(max(sa[0], sb[0]))
    ]


def _compare(left: Any, right: Any, operator: str, date_system: str) -> Any:
    for value in (left, right):
        if _error_kind(value) is not None:
            return value
    left, right = _date_serial(left, date_system), _date_serial(right, date_system)
    if left is None:
        left = "" if isinstance(right, str) else False if isinstance(right, bool) else 0
    if right is None:
        right = "" if isinstance(left, str) else False if isinstance(left, bool) else 0

    def ranked(value):
        if isinstance(value, bool):
            return 2, value
        if isinstance(value, (int, float)):
            return 0, value
        if isinstance(value, str):
            return 1, value.casefold()
        return None

    a, b = ranked(left), ranked(right)
    if a is None or b is None:
        return _error("Value")
    return {
        "=": a == b,
        "<>": a != b,
        "<": a < b,
        ">": a > b,
        "<=": a <= b,
        ">=": a >= b,
    }[operator]


def _concat(left: Any, right: Any, date_system: str) -> Any:
    for value in (left, right):
        if _error_kind(value) is not None:
            return value

    def text(value):
        value = _date_serial(value, date_system)
        if value is None:
            return ""
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, (int, float)):
            return format(value, ".15g")
        return str(value)

    result = text(left) + text(right)
    return result if len(result) <= 32767 else _error("Value")


def register_excel_functions(workbook, date_system: str = "1900") -> None:
    """Install workbook-local corrections before any formula evaluation."""
    # Excel formulas produce serial numbers, even when a native intermediate
    # retains a temporal tag (DATE(...)-serial can otherwise look like a date).
    workbook.set_temporal_egress("serial")
    workbook.register_function(
        "DATE",
        lambda year, month, day: _date_array(year, month, day, date_system),
        min_args=3,
        max_args=3,
        allow_override_builtin=True,
    )
    for part in ("DAY", "MONTH", "YEAR"):
        workbook.register_function(
            part,
            lambda value, _part=part: _date_part(value, _part, date_system),
            min_args=1,
            max_args=1,
            allow_override_builtin=True,
        )
    workbook.register_function(
        "SUMPRODUCT",
        _sumproduct,
        min_args=1,
        max_args=255,
        allow_override_builtin=True,
    )
    workbook.register_function(
        _CONCAT,
        lambda a, b: _binary_array(lambda x, y: _concat(x, y, date_system), a, b),
        min_args=2,
        max_args=2,
    )
    workbook.register_function(
        _COMPARE,
        lambda a, b, op: _binary_array(
            lambda x, y: _compare(x, y, op, date_system), a, b
        ),
        min_args=3,
        max_args=3,
    )


def read_name_bindings(data: bytes) -> dict[tuple[str | None, str], str]:
    """Read source expressions including constants and sheet-local shadowing."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = [sheet.attrib["name"] for sheet in root.findall("{*}sheets/{*}sheet")]
    result = NameBindings(sheets)
    for name in root.findall("{*}definedNames/{*}definedName"):
        label, expression = name.get("name"), name.text
        if not label or not expression:
            continue
        scope = None
        if name.get("localSheetId") is not None:
            try:
                scope = sheets[int(name.attrib["localSheetId"])]
            except (ValueError, IndexError):
                continue
        result[(scope, label.upper())] = expression
    return result


def rewrite_formula(
    formula: str,
    sheet: str,
    names: dict[tuple[str | None, str], str] | None = None,
) -> str:
    """Rewrite typed operators and resolvable names, retaining source elsewhere.

    AST leaves use the native renderer, which distinguishes an error token from
    a text literal spelling the same thing. LET/LAMBDA variables stay lexical.
    Unsafe or unparseable expressions are passed through, never guessed.
    """
    from linexcel.engine import is_too_deep

    expression = formula if formula.startswith("=") else "=" + formula
    if not names and not any(op in expression[1:] for op in "&=<>"):
        return formula
    if is_too_deep(expression):
        return formula
    try:
        root = fz.parse(expression)
    except Exception:
        return formula
    changed = False
    bindings = names or {}
    scopes = getattr(bindings, "sheet_names", {}) | {
        scope.casefold(): scope for scope, _ in bindings if scope is not None
    }
    visits = 0

    def render(node, scope, bound=frozenset(), resolving=frozenset()):
        nonlocal changed, visits
        visits += 1
        if visits > 10000:
            raise ValueError("bounded name/formula expansion")
        kind = node.node_type()
        children = node.children()
        if kind == "Reference" and bindings:
            reference = node.get_reference_string()
            qualified, bare = split_sheet_prefix(reference)
            if qualified is None and bare.upper() in bound:
                return node.to_formula().removeprefix("=")
            if qualified is not None and qualified.casefold() not in scopes:
                return node.to_formula().removeprefix("=")
            context = scopes.get((qualified or scope).casefold(), qualified or scope)
            key = (context, bare.upper())
            if key not in bindings:
                key = (None, bare.upper())
            if key in bindings and key not in resolving and len(resolving) < 32:
                expression = bindings[key]
                # Relative defined references require a base cell the package
                # does not record reliably. Leave these to the native engine.
                rect = parse_ref(expression)
                if rect is None or _absolute_reference(expression):
                    try:
                        if is_too_deep("=" + expression.lstrip("=")):
                            return node.to_formula().removeprefix("=")
                        definition = fz.parse("=" + expression.lstrip("="))
                        references = (str(ref) for ref in definition.walk_refs())
                        if any(
                            parse_ref(ref) is not None
                            and (
                                not _absolute_reference(ref)
                                or split_sheet_prefix(ref)[0] is None
                            )
                            for ref in references
                        ):
                            return node.to_formula().removeprefix("=")
                        replacement = render(
                            definition, key[0] or scope, bound, resolving | {key}
                        )
                    except Exception:
                        return node.to_formula().removeprefix("=")
                    changed = True
                    return "(" + replacement + ")"
            return node.to_formula().removeprefix("=")
        if kind == "BinaryOp":
            left, right = [render(c, scope, bound, resolving) for c in children]
            operator = node.get_operator()
            if operator == "&":
                changed = True
                return f"{_CONCAT}({left},{right})"
            if operator in _COMPARISONS:
                changed = True
                return f'{_COMPARE}({left},{right},"{operator}")'
            return f"({left}{operator}{right})"
        if kind == "Function":
            name = node.get_function_name()
            local = name.upper().removeprefix("_XLFN.")
            if local in {"LET", "LAMBDA"}:
                parts = []
                current = set(bound)
                for i, child in enumerate(children):
                    declaration = i < len(children) - 1 and (
                        local == "LAMBDA" or i % 2 == 0
                    )
                    if declaration:
                        parts.append(child.to_formula().removeprefix("="))
                        # LET value sees earlier bindings, not its own binding.
                        if local == "LAMBDA":
                            current.add(parts[-1].upper())
                    else:
                        parts.append(
                            render(child, scope, frozenset(current), resolving)
                        )
                        if local == "LET" and i < len(children) - 1:
                            current.add(parts[-2].upper())
                return f"{name}({','.join(parts)})"
            args = ",".join(render(c, scope, bound, resolving) for c in children)
            return f"{name}({args})"
        if kind == "UnaryOp":
            body = render(children[0], scope, bound, resolving)
            operator = node.get_operator()
            return f"({body})%" if operator == "%" else f"{operator}({body})"
        return node.to_formula().removeprefix("=")

    try:
        rendered = render(root, sheet)
    except (ValueError, RuntimeError, RecursionError):
        return formula
    rewritten = "=" + rendered
    # Several individually bounded named expressions can produce an unsafe
    # expanded AST. Never let expansion bypass the engine's native-stack guard.
    return rewritten if changed and not is_too_deep(rewritten) else formula


class CompatibleWorkbook:
    """Native evaluation with original formulas retained for lineage readers."""

    def __init__(self, native, names=None):
        self.native = native
        self.names = names or {}
        self.originals: dict[tuple[str, int, int], str] = {}

    def __getattr__(self, name):
        return getattr(self.native, name)

    def sheet(self, name):
        return _CompatibleSheet(self, name)

    def qualify_formula(self, formula, source_sheet):
        return rewrite_formula(formula, source_sheet, self.names)

    def set_formula(self, sheet, row, col, formula):
        rewritten = rewrite_formula(formula, sheet, self.names)
        self.native.set_formula(sheet, row, col, rewritten)
        if rewritten != formula:
            self.originals[(sheet, row, col)] = formula
        else:
            self.originals.pop((sheet, row, col), None)

    def get_formula(self, sheet, row, col):
        return self.originals.get((sheet, row, col)) or self.native.get_formula(
            sheet, row, col
        )

    def set_value(self, sheet, row, col, value):
        self.native.set_value(sheet, row, col, value)
        self.originals.pop((sheet, row, col), None)

    def set_formulas_batch(self, sheet, start_row, start_col, formulas):
        for r, row in enumerate(formulas, start_row):
            for c, formula in enumerate(row, start_col):
                if formula is not None:
                    self.set_formula(sheet, r, c, formula)


class _CompatibleSheet:
    def __init__(self, workbook, name):
        self.workbook = workbook
        self.name = name
        self.native = workbook.native.sheet(name)

    def __getattr__(self, name):
        return getattr(self.native, name)

    def set_formula(self, row, col, formula):
        return self.workbook.set_formula(self.name, row, col, formula)

    def set_value(self, row, col, value):
        return self.workbook.set_value(self.name, row, col, value)

    def get_formulas(self, rectangle):
        rows = self.native.get_formulas(rectangle)
        for r, row in enumerate(rows, rectangle.start_row):
            for c, formula in enumerate(row, rectangle.start_col):
                row[c - rectangle.start_col] = self.workbook.originals.get(
                    (self.name, r, c), formula
                )
        return rows
