"""Réécriture de formules via le tokenizer formualizer (Rust).

Deux usages :
- canonicalisation R1C1 relative à la cellule porteuse, pour détecter les
  formules « étirées » (recopiées) : deux cellules dont la forme R1C1 est
  identique portent la même logique ;
- qualification des références par un nom de feuille, pour évaluer une
  sous-expression dans une feuille de brouillon sans casser les références
  relatives.
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
    """Forme canonique R1C1 d'une formule, relative à (row, col).

    Les références A1 sont converties en offsets relatifs ; les noms définis
    et références structurées restent tels quels. Deux cellules issues d'une
    même recopie (étirement) produisent la même chaîne.
    """
    try:
        toks = _tokens(formula)
    except Exception:
        # Formule que le tokenizer ne comprend pas : la chaîne brute sert de clé.
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
    """Préfixe toutes les références non qualifiées par ``sheet``.

    Permet d'évaluer ``=SUM(A1:A10)`` (écrit dans Feuil1) depuis une feuille
    de brouillon : ``=SUM(Feuil1!A1:A10)``.
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
