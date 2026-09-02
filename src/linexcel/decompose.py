"""Step-by-step decomposition of a composite formula into evaluated steps.

Extracted mechanically from analyzer.py: the pure formula side of value
recovery. Each function walks the AST of a composite formula, renders a
subtree back to the text the engine can parse again, and evaluates a step on
its own in a scratch sheet. Everything here is driven by a passed-in
``resolver`` and returns plain dicts; nothing reaches back into the analyzer.

The scratch constants live here because the scratch sheet exists *for* step
decomposition — it is the engine sheet where guarded evaluations happen.
"""

from __future__ import annotations

import itertools
from typing import TYPE_CHECKING, Any

import formualizer as fz

from linexcel.external import parse_external_refs
from linexcel.refs import parse_ref_detailed
from linexcel.rewrite import qualify_sheet
from linexcel.values import _jsonable

if TYPE_CHECKING:  # type annotations only; never executed, so no import cycle
    from linexcel.analyzer import _ValueResolver
    from linexcel.refs import Rect

SCRATCH_SHEET = "__lineage_scratch__"
# Written into the scratch cell before each guarded evaluation: when the engine
# fails to compute an expression it silently keeps the previous cell value
# instead of raising, so an unchanged marker is how we detect that failure.
SCRATCH_SENTINEL = "__linexcel_no_value__"
GUARD_FUNCTIONS = {"IFERROR", "IFNA"}

MAX_STEPS_PER_FORMULA = 48

_STEP_KINDS = {"Function", "BinaryOp", "UnaryOp"}
#: Excel operator precedence, loosest first. An unknown operator scores 0 and
#: is therefore always parenthesized, which is the safe way to be wrong.
_PRECEDENCE = {
    "=": 1,
    "<": 1,
    ">": 1,
    "<=": 1,
    ">=": 1,
    "<>": 1,
    "&": 2,
    "+": 3,
    "-": 3,
    "*": 4,
    "/": 4,
    "^": 5,
}
#: Unary minus binds tighter than any binary operator: ``-2^2`` is 4 in Excel,
#: and the parser agrees — it reads ``-A1^2`` as ``(-A1)^2``.
_UNARY_PRECEDENCE = 6


def _collect_step_exprs(ast_dict: dict, *, skip_root: bool = False) -> list[str]:
    """First pass: collect every step expression ``_decompose`` will evaluate.

    The counter and ``_STEP_KINDS`` filter mirror ``_decompose`` exactly so
    the expressions batch-evaluated here are the same ones looked up later
    in ``_eval_raw``.  Order is pre-order (parent before children) — it does
    not match the post-order evaluation in ``_decompose`` but that is fine:
    the batch evaluates all at once and the cache is order-independent.

    ``skip_root`` leaves out the whole-formula expression, which the caller
    passes to ``_decompose`` as ``root_value`` instead. The counter still
    counts it, so both passes keep budgeting the same steps.
    """
    exprs: list[str] = []
    counter = itertools.count()

    def walk(node: dict, is_root: bool) -> None:
        ntype = node.get("node_type")
        if ntype not in _STEP_KINDS:
            return
        if next(counter) >= MAX_STEPS_PER_FORMULA:
            return
        if not (is_root and skip_root):
            exprs.append(_render_expr(node))
        if ntype == "Function":
            children = node.get("args", [])
        elif ntype == "BinaryOp":
            children = [node.get("left"), node.get("right")]
        else:
            children = [node.get("operand") or node.get("expr")]
        for c in children:
            if c:
                walk(c, False)

    walk(ast_dict, True)
    return exprs


def _decompose(
    ast_dict: dict,
    sheet: str,
    resolver: _ValueResolver,
    defined_names: dict[str, list[Rect]] | None = None,
    root_value: Any = None,
) -> dict | None:
    """Step tree: each function / operator becomes an evaluated step.

    ``root_value`` is the value the engine already computed for the cell
    itself. The root step *is* the whole formula, so re-evaluating it in the
    scratch sheet recomputes a value the workbook already holds — and a root
    such as ``SUM(Calculs!H2:H200001)`` makes the engine walk 200,000 formula
    cells again, which measured 29 s on one node. Pass it only when it came
    from the engine: a value read from the file, or an error-guarded fallback,
    is not what evaluating this expression yields.
    """
    counter = itertools.count()

    def expr_of(node: dict) -> str:
        return _render_expr(node)

    def walk(node: dict, depth: int) -> dict | None:
        ntype = node.get("node_type")
        if ntype not in _STEP_KINDS:
            return None
        if next(counter) >= MAX_STEPS_PER_FORMULA:
            return None
        expr = expr_of(node)
        if ntype == "Function":
            label = node.get("name", "?")
            children_ast = node.get("args", [])
        elif ntype == "BinaryOp":
            label = node.get("operator", "?")
            children_ast = [node.get("left"), node.get("right")]
        else:
            label = node.get("operator", "?")
            children_ast = [node.get("operand") or node.get("expr")]
        children_ast = [c for c in children_ast if c]

        inputs = []
        children = []
        for child in children_ast:
            sub = walk(child, depth + 1)
            if sub is not None:
                children.append(sub)
            else:
                ctype = child.get("node_type")
                if ctype == "Reference":
                    ref = child.get("reference", "?")
                    preview, date_text = _ref_preview(
                        resolver, ref, sheet, defined_names
                    )
                    entry: dict[str, Any] = {"ref": ref, "value": preview}
                    if date_text is not None:
                        entry["date"] = date_text
                    inputs.append(entry)
                elif ctype == "Literal":
                    inputs.append({"literal": child.get("value")})

        if depth == 0 and root_value is not None:
            value, evaluated = root_value, True
        else:
            value, evaluated = resolver.eval_expr(expr, sheet)
        return {
            "kind": ntype,
            "label": label,
            "expr": expr,
            "value": value,
            "evaluated": evaluated,
            "inputs": inputs,
            "children": children,
        }

    return walk(ast_dict, 0)


def _render_expr(node: dict) -> str:
    """Reconstruct the expression of an AST subtree (readable form)."""
    ntype = node.get("node_type")
    if ntype == "Function":
        args = ", ".join(_render_expr(a) for a in node.get("args", []))
        return f"{node.get('name', '?')}({args})"
    if ntype == "BinaryOp":
        operator = node.get("operator", "?")
        precedence = _PRECEDENCE.get(operator, 0)
        left_node, right_node = node.get("left", {}), node.get("right", {})
        # Rendered inline rather than through a helper: one Python frame per
        # AST level, and the stress workbook carries a 700-term chain of
        # additions — two frames per level exhausts the interpreter on it.
        left = _render_expr(left_node)
        if _needs_parens(left_node, precedence, right_side=False):
            left = f"({left})"
        right = _render_expr(right_node)
        if _needs_parens(right_node, precedence, right_side=True):
            right = f"({right})"
        return f"{left} {operator} {right}"
    if ntype == "UnaryOp":
        operator = node.get("operator", "?")
        operand = node.get("operand") or node.get("expr") or {}
        rendered = _render_expr(operand)
        if _needs_parens(operand, _UNARY_PRECEDENCE, right_side=True):
            rendered = f"({rendered})"
        # Percent is the one postfix operator: 25%, not %25.
        return f"{rendered}%" if operator == "%" else f"{operator}{rendered}"
    if ntype == "Reference":
        return str(node.get("reference", "?"))
    if ntype == "Literal":
        v = node.get("value")
        if isinstance(v, str):
            return '"' + v.replace('"', '""') + '"'
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    if ntype == "Array":
        return "{...}"
    if ntype == "Paren":
        inner = node.get("expr") or node.get("inner") or {}
        return f"({_render_expr(inner)})"
    return "?"


def _needs_parens(node: dict, parent_precedence: int, right_side: bool) -> bool:
    """Whether a child operand has to be parenthesized under its parent.

    The parser keeps grouping in the *shape* of the tree and drops the
    parentheses themselves, so rendering a subtree flat changes what the text
    means: ``=D2*(1-Rate)`` came back as ``D2 * 1 - Rate``. That is not only
    misread by a human — each step is evaluated by re-parsing its own rendered
    text, so the step reported 2470.06 for a cell holding 1976.208.

    A right operand of *equal* precedence is parenthesized too: ``A - (B - C)``
    and ``A - B - C`` are different sums, and spelling the grouping out keeps
    the text re-parsing to the very tree it was rendered from, whichever way
    the parser happens to associate.
    """
    if node.get("node_type") != "BinaryOp":
        return False
    precedence = _PRECEDENCE.get(node.get("operator", ""), 0)
    return precedence < parent_precedence or (
        right_side and precedence == parent_precedence
    )


def _scratch_eval(engine, expr: str, sheet: str) -> tuple[Any, bool]:
    """Evaluate an expression on its own in the scratch sheet.

    The cell is primed with a sentinel first: when the engine cannot compute
    an expression it leaves the previous value in place instead of reporting
    an error, and an unchanged sentinel is the only way to tell. The very
    first evaluation on a workbook holding a broken reference raises — the
    engine walks the whole dirty graph — while later ones stay isolated,
    hence the single retry.
    """
    try:
        qualified = qualify_sheet(expr, sheet)
    except Exception:
        return None, False
    for _ in range(2):
        try:
            engine.set_formula(SCRATCH_SHEET, 1, 1, f'="{SCRATCH_SENTINEL}"')
            if engine.evaluate_cell(SCRATCH_SHEET, 1, 1) != SCRATCH_SENTINEL:
                continue
            engine.set_formula(SCRATCH_SHEET, 1, 1, qualified)
            value = engine.evaluate_cell(SCRATCH_SHEET, 1, 1)
        except Exception:
            continue
        if value == SCRATCH_SENTINEL:
            return None, False
        return value, True
    return None, False


def _guard_fallback_expr(expr: str) -> str | None:
    """Fallback branch of a top-level IFERROR/IFNA, as Excel would show it."""
    # ponytail: only a top-level guard is recovered. With a nested one —
    # =IFERROR(SUM(IFERROR(NOSHEET!A1,0)),1) — the whole expression fails to
    # evaluate, so the outer fallback branch is taken (1) where Excel lets the
    # inner guard absorb the broken reference and returns 0. That gap is the
    # accepted ceiling of this recovery.
    try:
        ast_dict = fz.parse(expr).to_dict()
    except Exception:
        return None
    if not isinstance(ast_dict, dict) or ast_dict.get("node_type") != "Function":
        return None
    if str(ast_dict.get("name", "")).upper() not in GUARD_FUNCTIONS:
        return None
    args = ast_dict.get("args") or []
    if len(args) < 2:
        return None
    return "=" + _render_expr(args[1])


def _ref_preview(
    resolver: _ValueResolver,
    ref: str,
    sheet: str,
    defined_names: dict[str, list[Rect]] | None = None,
) -> tuple[Any, str | None]:
    """Preview of a referenced cell or range: ``(value, date_text)``."""
    # A reference into another workbook parses as nothing local; it still has a
    # value whenever that workbook was read, and the step is unreadable without.
    external = parse_external_refs(ref)
    if external:
        value, source = resolver.external_value(external[0])
        if source is not None:
            return _jsonable(value), None
    detail = parse_ref_detailed(ref, default_sheet=sheet)
    if detail is None:
        # may be a defined name: show the value of its target
        if defined_names:
            for name, rects in defined_names.items():
                if name.upper() == ref.upper() and rects:
                    rect = rects[0]
                    if rect.ncells == 1:
                        value, _, date_text = resolver.value(
                            rect.sheet or sheet, rect.r1, rect.c1
                        )
                        return value, date_text
                    return {"range": rect.to_a1(), "n": rect.ncells}, None
        return None, None
    rect = detail.rect
    if rect.ncells == 1:
        value, _, date_text = resolver.value(rect.sheet or sheet, rect.r1, rect.c1)
        return value, date_text
    return {"range": rect.to_a1(), "n": rect.ncells}, None
