"""Formula rewriting via the formualizer tokenizer (Rust).

Two uses:
- R1C1 canonicalization relative to the host cell, to detect "stretched"
  (copied) formulas: two cells with the same R1C1 form carry the same logic;
- qualifying references with a sheet name, to evaluate a sub-expression in a
  scratch sheet without breaking relative references.
"""

from __future__ import annotations

import re

import formualizer as fz

from linexcel.refs import parse_ref, quote_sheet, ref_to_r1c1, split_sheet_prefix

# An A1 cell inside a formula the tokenizer refused: optional sheet prefix,
# any $ anchors. The lookarounds keep it from biting into a longer word, so
# LOG10( and _xlfn.FOO2 are left alone.
_A1_IN_TEXT_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:(?:'(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
    r"(?![A-Za-z0-9_(])"
)


def _tokens(formula: str) -> list:
    if not formula.startswith("="):
        formula = "=" + formula
    return list(fz.tokenize(formula))


def _is_range_operand(token) -> bool:
    return str(token.token_type) == "Operand" and str(token.subtype) == "Range"


def _is_a1_ref(ref: str) -> bool:
    """Is this Range operand really an A1 reference?

    The tokenizer labels bare identifiers as Range operands: a LET/LAMBDA
    parameter (``x``) or a defined name short enough to look like a column
    (``TVA``). Handed to :func:`ref_to_r1c1` those become column offsets that
    move with the host cell, so a horizontally copied run stops sharing a key
    and explodes into one node per cell. Operands holding a ``:`` keep the old
    path: range endpoints (``A:A``, ``1:1``) and 3D refs (``S1:S3!A2``) are
    references even though :func:`parse_ref` declines them.
    """
    _sheet, body = split_sheet_prefix(ref)
    return ":" in body or parse_ref(body) is not None


def _structural_key(formula: str) -> str:
    """Grouping key for a formula the tokenizer could not read.

    Cell references collapse to ``@`` so two cells of the same copied run
    share the key, whitespace collapses (kept as one space: it is Excel's
    intersection operator), and case folds.

    # ponytail: regex normalization is strictly weaker than R1C1. It sees no
    # structure, so two *different* runs that differ only in what the regex
    # erases collide into one node, and a run whose cells differ by a whole
    # column ref (A:A vs B:B) still splits. Both are acceptable here: the
    # worst case merges or keeps apart nodes, it never corrupts a formula.
    """
    normalized = _A1_IN_TEXT_RE.sub("@", formula)
    return re.sub(r"\s+", " ", normalized).strip().lower()


def canonical_r1c1(formula: str, row: int, col: int) -> str:
    """Canonical R1C1 form of a formula, relative to (row, col).

    A1 references are converted to relative offsets; defined names and
    structured references stay as-is. Two cells from the same copy (stretch)
    produce the same string.
    """
    try:
        toks = _tokens(formula)
    except Exception:
        # Formula the tokenizer can't understand: fall back to a structural
        # key, which still groups the cells of one copied run together.
        return _structural_key(formula)
    out: list[str] = []
    for t in toks:
        v = t.value
        if _is_range_operand(t) and _is_a1_ref(v):
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
