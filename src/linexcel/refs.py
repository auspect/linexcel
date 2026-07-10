"""Utilitaires de références Excel : A1 ↔ (ligne, colonne), R1C1, plages.

Toutes les conversions sont indépendantes du moteur de calcul afin de
rester testables unitairement.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

MAX_ROW = 1_048_576
MAX_COL = 16_384

_CELL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_COL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^(\$?)(\d+)$")

# Référence complète éventuellement préfixée d'une feuille :
#   'Ma Feuille'!A1:B2   Feuil1!$C$3   A1   A:B   1:4
_SHEET_PREFIX_RE = re.compile(
    r"^(?:(?P<q>'(?:[^']|'')+')|(?P<p>[^'!()+\-*/^&=<>,; ]+))!"
)


def col_to_num(col: str) -> int:
    """Convertit une colonne alphabétique (A, B, …, XFD) en numéro 1-indexé."""
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n: int) -> str:
    """Convertit un numéro de colonne 1-indexé en lettres."""
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


@dataclass(frozen=True)
class Rect:
    """Rectangle de cellules (bornes incluses), avec feuille optionnelle."""

    sheet: str | None
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def ncells(self) -> int:
        return (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)

    def clipped(self, max_row: int, max_col: int) -> Rect | None:
        """Restreint la plage aux dimensions utilisées de la feuille."""
        r2 = min(self.r2, max(max_row, 1))
        c2 = min(self.c2, max(max_col, 1))
        if r2 < self.r1 or c2 < self.c1:
            return None
        return Rect(self.sheet, self.r1, self.c1, r2, c2)

    def intersects(self, other: Rect) -> bool:
        return not (
            self.r2 < other.r1
            or other.r2 < self.r1
            or self.c2 < other.c1
            or other.c2 < self.c1
        )

    def to_a1(self) -> str:
        start = f"{num_to_col(self.c1)}{self.r1}"
        end = f"{num_to_col(self.c2)}{self.r2}"
        addr = start if start == end else f"{start}:{end}"
        return f"{quote_sheet(self.sheet)}!{addr}" if self.sheet else addr


def quote_sheet(sheet: str) -> str:
    """Cite un nom de feuille si nécessaire pour l'insérer dans une formule."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def split_sheet_prefix(ref: str) -> tuple[str | None, str]:
    """Sépare le préfixe de feuille d'une référence ('Feuil 1'!A1 → (Feuil 1, A1))."""
    m = _SHEET_PREFIX_RE.match(ref)
    if not m:
        return None, ref
    if m.group("q"):
        sheet = m.group("q")[1:-1].replace("''", "'")
    else:
        sheet = m.group("p")
    return sheet, ref[m.end() :]


def parse_ref(ref: str, default_sheet: str | None = None) -> Rect | None:
    """Analyse une référence A1 (cellule, plage, colonnes ou lignes entières).

    Retourne ``None`` si la chaîne n'est pas une référence A1 valide
    (nom défini, référence structurée de tableau, …).
    """
    sheet, body = split_sheet_prefix(ref)
    if sheet is None:
        sheet = default_sheet
    elif ":" in sheet:
        # Référence 3D (Feuil1:Feuil3!A1) : hors périmètre du graphe cellule.
        return None
    body = body.strip()
    if ":" in body:
        left, _, right = body.partition(":")
        p1 = _parse_endpoint(left)
        p2 = _parse_endpoint(right)
        if p1 is None or p2 is None:
            return None
        (r1, c1), (r2, c2) = p1, p2
        rect = Rect(sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
        return normalize_whole_ranges(rect)
    m = _CELL_RE.match(body)
    if not m:
        return None
    row = int(m.group(4))
    col = col_to_num(m.group(2))
    if row < 1 or row > MAX_ROW or col > MAX_COL:
        return None
    return Rect(sheet, row, col, row, col)


def _parse_endpoint(part: str) -> tuple[int, int] | None:
    """Borne d'une plage : cellule, colonne entière ou ligne entière."""
    part = part.strip()
    m = _CELL_RE.match(part)
    if m:
        return int(m.group(4)), col_to_num(m.group(2))
    m = _COL_RE.match(part)
    if m:
        col = col_to_num(m.group(2))
        if col > MAX_COL:
            return None
        # Colonne entière : les lignes seront bornées par _COL/_ROW jumelé.
        return -1, col
    m = _ROW_RE.match(part)
    if m:
        row = int(m.group(2))
        if row < 1 or row > MAX_ROW:
            return None
        return row, -1
    return None


def normalize_whole_ranges(rect: Rect) -> Rect:
    """Remplace les marqueurs -1 (colonne/ligne entière) par les bornes max."""
    r1 = 1 if rect.r1 == -1 else rect.r1
    r2 = MAX_ROW if rect.r2 == -1 else rect.r2
    c1 = 1 if rect.c1 == -1 else rect.c1
    c2 = MAX_COL if rect.c2 == -1 else rect.c2
    return Rect(rect.sheet, r1, c1, r2, c2)


@dataclass(frozen=True)
class RefDetail:
    """Référence analysée avec ses ancres $ (pour l'étirement des groupes)."""

    rect: Rect
    row1_abs: bool
    col1_abs: bool
    row2_abs: bool
    col2_abs: bool


def parse_ref_detailed(ref: str, default_sheet: str | None = None) -> RefDetail | None:
    """Comme :func:`parse_ref`, mais conserve les ancres absolues de chaque borne."""
    sheet, body = split_sheet_prefix(ref)
    if sheet is None:
        sheet = default_sheet
    elif ":" in sheet:
        return None
    body = body.strip()
    parts = body.split(":")
    if len(parts) > 2:
        return None
    ends = []
    for part in parts:
        e = _parse_endpoint_detailed(part.strip())
        if e is None:
            return None
        ends.append(e)
    if len(ends) == 1:
        (row, col, row_abs, col_abs) = ends[0]
        if row == -1 or col == -1:
            return None  # une colonne/ligne entière seule n'est pas une cellule
        rect = Rect(sheet, row, col, row, col)
        return RefDetail(rect, row_abs, col_abs, row_abs, col_abs)
    (r1, c1, r1a, c1a), (r2, c2, r2a, c2a) = ends
    if (r1, c1) > (r2, c2):
        (r1, c1, r1a, c1a), (r2, c2, r2a, c2a) = (r2, c2, r2a, c2a), (r1, c1, r1a, c1a)
    # Les axes « entiers » (A:A, 1:1) sont figés : ils ne s'étirent pas.
    if r1 == -1:
        r1, r2, r1a, r2a = 1, MAX_ROW, True, True
    if c1 == -1:
        c1, c2, c1a, c2a = 1, MAX_COL, True, True
    return RefDetail(Rect(sheet, r1, c1, r2, c2), r1a, c1a, r2a, c2a)


def _parse_endpoint_detailed(part: str) -> tuple[int, int, bool, bool] | None:
    m = _CELL_RE.match(part)
    if m:
        row, col = int(m.group(4)), col_to_num(m.group(2))
        if row < 1 or row > MAX_ROW or col > MAX_COL:
            return None
        return row, col, m.group(3) == "$", m.group(1) == "$"
    m = _COL_RE.match(part)
    if m:
        col = col_to_num(m.group(2))
        if col > MAX_COL:
            return None
        return -1, col, True, m.group(1) == "$"
    m = _ROW_RE.match(part)
    if m:
        row = int(m.group(2))
        if row < 1 or row > MAX_ROW:
            return None
        return row, -1, m.group(1) == "$", True
    return None


def stretch_ref(
    detail: RefDetail,
    rep_row: int,
    rep_col: int,
    rows_span: tuple[int, int],
    cols_span: tuple[int, int],
) -> Rect:
    """Étend la référence d'une cellule représentative à tout un groupe étiré.

    ``rows_span``/``cols_span`` sont les bornes (min, max) des cellules membres
    du groupe. Les bornes relatives suivent le déplacement, les bornes ancrées
    ($) restent fixes.
    """
    rect = detail.rect
    r1 = rect.r1 if detail.row1_abs else rect.r1 + (rows_span[0] - rep_row)
    r2 = rect.r2 if detail.row2_abs else rect.r2 + (rows_span[1] - rep_row)
    c1 = rect.c1 if detail.col1_abs else rect.c1 + (cols_span[0] - rep_col)
    c2 = rect.c2 if detail.col2_abs else rect.c2 + (cols_span[1] - rep_col)
    r1, r2 = max(1, min(r1, r2)), min(MAX_ROW, max(r1, r2))
    c1, c2 = max(1, min(c1, c2)), min(MAX_COL, max(c1, c2))
    return Rect(rect.sheet, r1, c1, r2, c2)


def cell_to_r1c1(cell: str, base_row: int, base_col: int) -> str | None:
    """Convertit une référence de cellule A1 en R1C1 relatif à (base_row, base_col)."""
    m = _CELL_RE.match(cell)
    if m:
        abs_col, col, abs_row, row = (
            m.group(1) == "$",
            col_to_num(m.group(2)),
            m.group(3) == "$",
            int(m.group(4)),
        )
        r = f"R{row}" if abs_row else _rel("R", row - base_row)
        c = f"C{col}" if abs_col else _rel("C", col - base_col)
        return r + c
    m = _COL_RE.match(cell)
    if m:
        abs_col, col = m.group(1) == "$", col_to_num(m.group(2))
        return f"C{col}" if abs_col else _rel("C", col - base_col)
    m = _ROW_RE.match(cell)
    if m:
        abs_row, row = m.group(1) == "$", int(m.group(2))
        return f"R{row}" if abs_row else _rel("R", row - base_row)
    return None


def _rel(axis: str, delta: int) -> str:
    return axis if delta == 0 else f"{axis}[{delta}]"


def ref_to_r1c1(ref: str, base_row: int, base_col: int) -> str | None:
    """Convertit une référence complète (avec feuille et ':') en forme R1C1."""
    sheet, body = split_sheet_prefix(ref)
    parts = body.split(":")
    if len(parts) > 2:
        return None
    converted = []
    for part in parts:
        c = cell_to_r1c1(part.strip(), base_row, base_col)
        if c is None:
            return None
        converted.append(c)
    out = ":".join(converted)
    if sheet is not None:
        out = f"{quote_sheet(sheet)}!{out}"
    return out
