"""What a cell value *is*, once something has read it.

The engine, the file's own cache and a linked workbook each hand values back
in their own shape — a Rust error object, a float that is really a date, a
string that is really an error code. This module is the one place that decides
what such a thing means: whether it counts as computed, whether two readings
of it agree, and how it is written down for the report.

Kept apart from the reading (:mod:`linexcel.loader`) and from the reasoning
(:mod:`linexcel.analyzer`) because both need these answers and neither owns
them.
"""

from __future__ import annotations

import datetime
import re
from typing import Any

EXCEL_EPOCH_1900 = datetime.datetime(1899, 12, 30)
EXCEL_EPOCH_1904 = datetime.datetime(1904, 1, 1)
# Serials 1..59 (Jan/Feb 1900) sit before Excel's phantom 1900-02-29, so their
# real epoch is 1899-12-31: serial 1 → 1900-01-01. Serial 60 is the phantom
# day itself and matches no real date. From 61 on, the 1899-12-30 epoch
# already absorbs the phantom day.
EPOCH_EARLY_1900 = datetime.datetime(1899, 12, 31)
#: Engine error kinds → the text a spreadsheet shows for them. These are values
#: a cell genuinely holds: openpyxl reads the stored one back as this same text,
#: so the graph carries the text on both sides and they compare directly.
ERROR_KIND_TEXT = {
    "Div": "#DIV/0!",
    "Na": "#N/A",
    "Name": "#NAME?",
    "Num": "#NUM!",
    "Null": "#NULL!",
    "Ref": "#REF!",
    "Value": "#VALUE!",
    "Spill": "#SPILL!",
    "Calc": "#CALC!",
}
EXCEL_ERRORS = frozenset(ERROR_KIND_TEXT.values())
#: Error kinds that are *not* a cell value: the engine reporting that it could
#: not compute, rather than a result. ``NImpl`` is a function or operator it does
#: not implement — the range intersection in ``=SUM(D2:D10 D5:D20)``, say — and
#: ``Circ`` a reference cycle, which Excel resolves by convention rather than by
#: computing. Neither may be shown as "the value linexcel recalculated", because
#: linexcel recalculated nothing; the cell falls back to what the file stores.
UNCOMPUTED_ERROR_KINDS = frozenset({"NImpl", "Cancelled", "Circ"})


def serial_to_date_text(serial: Any, epoch_1904: bool = False) -> str | None:
    """Excel serial number → ``YYYY-MM-DD``, or None if it is not a date."""
    if isinstance(serial, bool) or not isinstance(serial, (int, float)):
        return None
    days = float(serial)
    if days != days or days in (float("inf"), float("-inf")):
        return None
    if epoch_1904:
        if days < 0:
            return None
        base = EXCEL_EPOCH_1904
    else:
        if days < 1 or days == 60:
            return None
        base = EPOCH_EARLY_1900 if days < 60 else EXCEL_EPOCH_1900
    try:
        return (base + datetime.timedelta(days=days)).date().isoformat()
    except (OverflowError, ValueError):
        return None


def _date_text_of(value: Any) -> str | None:
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    return None


def _error_kind(value: Any) -> str | None:
    """Kind of an engine error value, or ``None`` when it is not one.

    The engine reports errors as a plain ``{"type": "Error", "kind": ...}``
    dict, which stringifies to Python source rather than to anything a reader
    of a spreadsheet recognises — hence every path out of the engine going
    through here first.
    """
    if isinstance(value, dict) and value.get("type") == "Error":
        kind = value.get("kind")
        return kind if isinstance(kind, str) else ""
    return None


def _excel_error_text(value: Any) -> str | None:
    """Spreadsheet text of an engine error value that a cell can hold."""
    kind = _error_kind(value)
    return None if kind is None else ERROR_KIND_TEXT.get(kind)


def _is_uncomputed(value: Any) -> bool:
    """True when the engine reported a limitation instead of a result.

    An unknown kind counts too: a kind this table does not know is one whose
    spreadsheet meaning we cannot vouch for, and inventing a value for it is
    worse than admitting the cell was not recalculated.
    """
    kind = _error_kind(value)
    if kind is None:
        return False
    return kind in UNCOMPUTED_ERROR_KINDS or kind not in ERROR_KIND_TEXT


def _values_differ(raw: Any, cached: Any, date_text: str | None) -> bool:
    """True when a recalculated value contradicts the one stored in the file."""
    cached_date = _date_text_of(cached)
    if cached_date is not None:
        return date_text is not None and date_text != cached_date
    # An error is a value a cell genuinely holds, and openpyxl reads the stored
    # one back as the same text Excel shows — so the two are directly
    # comparable, and a recalculated #DIV/0! over a stored #DIV/0! is agreement,
    # not a disagreement nobody can explain.
    error_text = _excel_error_text(raw)
    if error_text is not None or isinstance(cached, str) and cached in EXCEL_ERRORS:
        return error_text != cached
    if isinstance(raw, bool) or isinstance(cached, bool):
        return False
    if isinstance(raw, (int, float)) and isinstance(cached, (int, float)):
        return abs(float(raw) - float(cached)) > 1e-9
    return False


def _fmt_value(value: Any) -> str:
    return _date_text_of(value) or _excel_error_text(value) or str(value)


def _jsonable(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        is_nan_or_inf = value != value or value in (float("inf"), float("-inf"))
        if isinstance(value, float) and is_nan_or_inf:
            return str(value)
        return value
    error_text = _excel_error_text(value)
    if error_text is not None:
        return error_text
    return str(value)


# ──────────────────────────────────────────────
# Two readings of the same cell, written differently
# ──────────────────────────────────────────────

#: A number as some spreadsheet somewhere writes it: digits, with runs
#: separated by a group mark — space, non-breaking space, narrow no-break
#: space, apostrophe, dot or comma — and possibly a decimal mark.
_NUMBER_RE = re.compile(r"\d[\d   '.,]*\d|\d")
#: Marks a spreadsheet may use to group digits. The dot and the comma are in
#: here too: which of the two groups and which one separates the decimal part
#: is exactly what changes from one regional setting to the next.
_GROUP_MARKS = " \u00a0\u202f'.,"
#: Digits grouped the way every convention groups them — three at a time.
_GROUPED_RE = re.compile(r"^\d{1,3}(?:[ \u00a0\u202f'.,]\d{3})*$|^\d+$")


def _read_number(text: str) -> set[float]:
    """Every value ``text`` could be, read as a number.

    A set, because some spellings are genuinely ambiguous: ``1,234`` is a
    thousand-and-something to an English reader and one-point-two-three-four
    to a French one, and nothing in the string says which. Returning both is
    what lets the caller ask "could these two be the same number?" without
    pretending to know a locale the file never recorded.

    Grouping is held to the convention it comes from — groups of exactly
    three digits. Without that, ``1 2`` reads as twelve, and two genuinely
    different texts would be called one value written two ways.
    """
    readings: set[float] = set()
    for decimal_mark in (".", ","):
        parts = text.split(decimal_mark)
        if len(parts) > 2:
            continue
        whole, fraction = parts[0], parts[1] if len(parts) == 2 else ""
        if fraction and not fraction.isdigit():
            continue
        if not _GROUPED_RE.match(whole):
            continue
        digits = whole
        for mark in _GROUP_MARKS:
            digits = digits.replace(mark, "")
        try:
            readings.add(float(f"{digits}.{fraction}" if fraction else digits))
        except ValueError:
            continue
    return readings


def _separators_only(left: str, right: str) -> bool:
    """True when two strings differ in nothing but how numbers are written.

    ``6.7 €`` against ``6,7 €``: linexcel computed with a dot because that is
    what the engine does, and the file stores a comma because that is what the
    machine that saved it does. The cell is right both times, and calling it a
    disagreement blames the workbook for a regional setting.

    The text around the numbers must match exactly — this is not a fuzzy
    comparison. Only the numbers are read leniently, and only in the sense of
    "is there a reading under which these are the same value".
    """
    left_numbers = list(_NUMBER_RE.finditer(left))
    right_numbers = list(_NUMBER_RE.finditer(right))
    if len(left_numbers) != len(right_numbers):
        return False
    if _NUMBER_RE.sub("#", left) != _NUMBER_RE.sub("#", right):
        return False
    if not left_numbers:
        return False  # no numbers at all: the strings simply differ
    return all(
        _read_number(a.group()) & _read_number(b.group())
        for a, b in zip(left_numbers, right_numbers)
    )


def readings_agree(recalculated: Any, stored: Any, date_text: str | None) -> str:
    """How the two readings of one cell relate: ``same``, ``format`` or ``differ``.

    One rule, in one place. The report used to hold two: Python compared
    values for the warnings and returned "no difference" for every pair of
    strings — missing a recalculated ``non`` over a stored ``oui`` — while the
    viewer compared the rendered text and called any difference a
    disagreement, which is what made a French workbook light up red over
    ``6,7 €``.

    ``format`` is the middle answer neither of them had: the same value, spelt
    with the separators of whatever saved the file. Both readings are still
    shown; only the verdict softens.
    """
    if _values_differ(recalculated, stored, date_text):
        return "differ"
    left, right = _fmt_value(recalculated), _fmt_value(stored)
    if left == right:
        return "same"
    if _separators_only(left, right):
        return "format"
    return (
        "differ"
        if isinstance(recalculated, str) and isinstance(stored, str)
        else "same"
    )
