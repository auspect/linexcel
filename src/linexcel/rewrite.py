"""Formula rewriting via the formualizer tokenizer (Rust).

Two uses:
- R1C1 canonicalization relative to the host cell, to detect "stretched"
  (copied) formulas: two cells with the same R1C1 form carry the same logic;
- qualifying references with a sheet name, to evaluate a sub-expression in a
  scratch sheet without breaking relative references.
"""

from __future__ import annotations

import formualizer as fz

from linexcel.refs import parse_ref, quote_sheet, ref_to_r1c1, split_sheet_prefix


def _tokens(formula: str) -> list:
    if not formula.startswith("="):
        formula = "=" + formula
    return list(fz.tokenize(formula))


def _is_range_operand(token) -> bool:
    return str(token.token_type) == "Operand" and str(token.subtype) == "Range"


def canonical_r1c1(formula: str, row: int, col: int) -> str:
    """Canonical R1C1 form of a formula, relative to (row, col).

    A1 references are converted to relative offsets; defined names and
    structured references stay as-is. Two cells from the same copy (stretch)
    produce the same string.
    """
    try:
        toks = _tokens(formula)
    except Exception:
        # Formula the tokenizer can't understand: raw string serves as key.
        return formula
    out: list[str] = []
    for t in toks:
        v = t.value
        if _is_range_operand(t):
            conv = ref_to_r1c1(v, row, col)
            if conv is not None:
                v = conv
        out.append(v)
    return "".join(out)


def qualify_sheet(formula: str, sheet: str) -> str:
    """Prefix all unqualified references with ``sheet``.

    Allows evaluating ``=SUM(A1:A10)`` (written in Sheet1) from a scratch
    sheet: ``=SUM(Sheet1!A1:A10)``.
    """
    try:
        toks = _tokens(formula)
    except Exception:
        return formula
    out: list[str] = []
    for t in toks:
        v = t.value
        if _is_range_operand(t):
            existing_sheet, _body = split_sheet_prefix(v)
            if existing_sheet is None and parse_ref(v) is not None:
                v = f"{quote_sheet(sheet)}!{v}"
        out.append(v)
    return "=" + "".join(out)
