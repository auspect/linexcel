"""Detect ordinary formulas whose Excel result relies on implicit intersection.

An ordinary ``<f>`` formula (no ``t="array"``) is saved with legacy semantics:
where Excel expects one value but receives a multi-cell reference, it
intersects that reference with the formula's own row or column. ``=A1:A5*10``
in row 3 reads A3; ``=SUM(A1:A3*2)`` in row 8 is ``#VALUE!``. formualizer
evaluates such formulas with dynamic-array semantics instead (see
``docs/critical_review.md``, F05), so its value is not Excel's.

This module only *detects* the formulas concerned; it never rewrites them.
The rule is deliberately conservative. A multi-cell reference is concerned
when it sits in a scalar position:

* the whole formula (``=A:A``);
* an operand of an arithmetic, comparison, concatenation or unary operator
  (``=A1:A5*10``, ``=SUM(A1:A3*2)``);
* a direct argument of a function, unless that argument is known to accept a
  range (``SUM(A1:A3)``, ``VLOOKUP(x,D:F,2,0)``, ``COUNTIF(A:A,"x")``).

Everything below ``SUMPRODUCT`` is evaluated as an array by Excel, so it is
never concerned. Unknown functions do not accept ranges here: a false alarm
only turns a calculation into a named refusal, never into a false value.

Known gaps: a multi-cell reference returned by a function (``OFFSET``,
``INDEX(...,0)``, ``CHOOSE``, ``IF``) is not detected unless the index has
rewritten it into a static reference (``evaluationFormula``).
"""

from __future__ import annotations

from typing import Any

from linexcel.refs import parse_ref, split_sheet_prefix

#: Functions where every argument accepts a reference or an array.
_ALL_ARGUMENTS = frozenset(
    """
    SUM AVERAGE AVERAGEA MIN MAX MINA MAXA COUNT COUNTA COUNTBLANK PRODUCT
    SUMSQ MEDIAN MODE MODE.SNGL MODE.MULT STDEV STDEVP STDEV.S STDEV.P STDEVA
    STDEVPA VAR VARP VAR.S VAR.P VARA VARPA AVEDEV DEVSQ GEOMEAN HARMEAN KURT
    SKEW AND OR XOR CONCAT TEXTJOIN
    """.split()
)

#: Functions where only some argument positions (0-based) accept a reference.
_SOME_ARGUMENTS: dict[str, set[int]] = {
    "VLOOKUP": {1},
    "HLOOKUP": {1},
    "LOOKUP": {1, 2},
    "MATCH": {1},
    "XLOOKUP": {1, 2},
    "XMATCH": {1},
    "INDEX": {0},
    "OFFSET": {0},
    "ROW": {0},
    "COLUMN": {0},
    "ROWS": {0},
    "COLUMNS": {0},
    "AREAS": {0},
    "ISREF": {0},
    "CELL": {1},
    "COUNTIF": {0},
    "SUMIF": {0, 2},
    "AVERAGEIF": {0, 2},
    "RANK": {1},
    "RANK.EQ": {1},
    "RANK.AVG": {1},
    "LARGE": {0},
    "SMALL": {0},
    "PERCENTILE": {0},
    "PERCENTILE.INC": {0},
    "PERCENTILE.EXC": {0},
    "PERCENTRANK": {0},
    "PERCENTRANK.INC": {0},
    "PERCENTRANK.EXC": {0},
    "QUARTILE": {0},
    "QUARTILE.INC": {0},
    "QUARTILE.EXC": {0},
    "CORREL": {0, 1},
    "PEARSON": {0, 1},
    "COVAR": {0, 1},
    "COVARIANCE.P": {0, 1},
    "COVARIANCE.S": {0, 1},
    "SLOPE": {0, 1},
    "INTERCEPT": {0, 1},
    "RSQ": {0, 1},
    "STEYX": {0, 1},
    "FORECAST": {1, 2},
    "FORECAST.LINEAR": {1, 2},
    "IRR": {0},
    "MIRR": {0},
    "XNPV": {1, 2},
    "XIRR": {0, 1},
    "SUMX2MY2": {0, 1},
    "SUMX2PY2": {0, 1},
    "SUMXMY2": {0, 1},
}

#: Functions accepting references from a given position onwards.
_FROM_ARGUMENT = {"SUBTOTAL": 1, "AGGREGATE": 2, "NPV": 1}

#: Functions whose arguments Excel evaluates as arrays in a normal cell.
_ARRAY_CONTEXT = frozenset({"SUMPRODUCT"})

#: Range-building operators: their operands form a reference, not values.
_REFERENCE_OPERATORS = frozenset({":", " ", ","})

_PREFIXES = ("_XLFN._XLWS.", "_XLFN.", "_XLWS.")


def _function(name: str) -> str:
    name = name.upper()
    for prefix in _PREFIXES:
        if name.startswith(prefix):
            return name[len(prefix) :]
    return name


def _accepts_range(function: str, index: int) -> bool:
    if function in _ALL_ARGUMENTS:
        return True
    if function in _FROM_ARGUMENT:
        return index >= _FROM_ARGUMENT[function]
    if function in {"COUNTIFS"}:
        return index % 2 == 0
    if function in {"SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"}:
        return index == 0 or index % 2 == 1
    return index in _SOME_ARGUMENTS.get(function, ())


def _local(name: str) -> str:
    upper = name.upper()
    return upper[6:] if upper.startswith("_XLPM.") else upper


class _Detector:
    def __init__(self):
        self.found: str | None = None
        self.budget = 10000

    @staticmethod
    def multi(text: str, bound: dict[str, bool]) -> bool:
        """Whether one resolved reference may denote more than one cell."""
        prefix, body = split_sheet_prefix(text)
        if prefix is None and _local(body) in bound:
            return bound[_local(body)]
        if text.endswith("#"):
            return True
        if "[" in text:
            lowered = text.lower()
            return not ("[@" in lowered or "#this row" in lowered)
        rect = parse_ref(text)
        return rect is not None and rect.ncells > 1

    def walk(self, item: Any, context: str, bound: dict[str, bool]) -> None:
        if self.found is not None or not isinstance(item, dict):
            return
        self.budget -= 1
        if self.budget < 0:
            self.found = "(formula too large to classify)"
            return
        kind = item.get("node_type")
        if kind == "Reference":
            text = str(item.get("reference") or "")
            if context == "scalar" and self.multi(text, bound):
                self.found = text
        elif kind == "BinaryOp":
            operator = item.get("operator")
            if operator in _REFERENCE_OPERATORS:
                if context == "scalar":
                    self.found = _render(item)
                    return
                inner = "range"
            else:
                inner = "array" if context == "array" else "scalar"
            self.walk(item.get("left"), inner, bound)
            self.walk(item.get("right"), inner, bound)
        elif kind == "UnaryOp":
            self.walk(
                item.get("operand"), "array" if context == "array" else "scalar", bound
            )
        elif kind == "Function":
            self.function(item, context, bound)

    def function(self, item: dict, context: str, bound: dict[str, bool]) -> None:
        name = _function(str(item.get("name") or ""))
        args = item.get("args") or []
        if name == "LET":
            local = dict(bound)
            for index, arg in enumerate(args):
                last = index == len(args) - 1
                if not last and index % 2 == 0:
                    continue
                if last:
                    self.walk(arg, context, local)
                    continue
                # A binding value may hold a range; the variable carries it.
                self.walk(arg, "array" if context == "array" else "range", local)
                declaration = args[index - 1]
                if isinstance(declaration, dict):
                    variable = _local(str(declaration.get("reference") or ""))
                    local[variable] = (
                        isinstance(arg, dict)
                        and arg.get("node_type") == "Reference"
                        and self.multi(str(arg.get("reference") or ""), local)
                    )
            return
        if name == "LAMBDA":
            local = dict(bound)
            for arg in args[:-1]:
                if isinstance(arg, dict):
                    local[_local(str(arg.get("reference") or ""))] = False
            if args:
                self.walk(args[-1], "scalar", local)
            return
        array = context == "array" or name in _ARRAY_CONTEXT
        for index, arg in enumerate(args):
            if array:
                inner = "array"
            elif _accepts_range(name, index):
                inner = "range"
            else:
                inner = "scalar"
            self.walk(arg, inner, bound)


def _render(item: dict) -> str:
    if item.get("node_type") == "Reference":
        return str(item.get("reference") or "")
    if item.get("node_type") == "BinaryOp":
        return (
            _render(item.get("left") or {})
            + str(item.get("operator"))
            + _render(item.get("right") or {})
        )
    return "…"


def _parse(formula: str) -> dict | None:
    import formualizer as fz

    from linexcel.engine import is_too_deep

    expression = formula if formula.startswith("=") else "=" + formula
    if is_too_deep(expression):
        return None
    try:
        return fz.parse(expression).to_dict()
    except Exception:
        return None


def implicit_intersection_reference(formula: str) -> str | None:
    """Return the first multi-cell reference read in a scalar position.

    ``formula`` is the expression given to the engine, after linexcel has
    replaced defined names and table references by what they resolve to. The
    caller decides which cells are ordinary formulas: declared array formulas
    (``t="array"``) are never concerned. ``None`` means every multi-cell
    reference sits where Excel accepts a range. A formula the parser cannot
    read is reported as concerned, so that no unverified value is presented
    as Excel's.
    """
    tree = _parse(formula)
    if tree is None:
        return "(formula not parsed)"
    detector = _Detector()
    detector.walk(tree, "scalar", {})
    return detector.found
