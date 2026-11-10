"""Read a search query the way a reader types it: a displayed number or an address.

Excel shows 967.5 for a saved 967.538 and 6600 for 6600.000000000001. A reader
searches what the sheet shows, so a number matches every saved value that
rounds to it at the number of decimals typed.
"""

from __future__ import annotations

import math
import re

_GROUP = re.compile(r"(?<=\d)[   '’](?=\d{3}(?!\d))")
_PLAIN = re.compile(r"[+-]?\d+(?:\.(\d+))?")
_ADDRESS = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|([^'!\s][^!]*))!)?"
    r"\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})"
)
MAX_ROW, MAX_COLUMN = 1_048_576, 16_384


def _range(text: str) -> tuple[float, float] | None:
    match = _PLAIN.fullmatch(text)
    if not match:
        return None
    value = float(text)
    if not math.isfinite(value):
        return None
    half = 0.5 * 10 ** -len(match.group(1) or "")
    return value - half, value + half


def _grouped(text: str, separator: str) -> bool:
    """True for digits in groups of three, as in 1,455,900."""
    return bool(
        re.fullmatch(rf"[+-]?\d{{1,3}}(?:{re.escape(separator)}\d{{3}})+", text)
    )


def _spellings(text: str) -> list[str]:
    """Plain decimal spellings of a number typed with separators."""
    comma, dot = text.count(","), text.count(".")
    if comma and dot:
        # The last separator is the decimal one: 1,455,900.5 or 1.455.900,5.
        if text.rfind(",") > text.rfind("."):
            whole, _, fraction = text.rpartition(",")
            grouped = dot == 1 or _grouped(whole, ".")
            return [f"{whole.replace('.', '')}.{fraction}"] if grouped else []
        whole, _, fraction = text.rpartition(".")
        grouped = comma == 1 or _grouped(whole, ",")
        return [f"{whole.replace(',', '')}.{fraction}"] if grouped else []
    if comma > 1:
        return [text.replace(",", "")] if _grouped(text, ",") else []
    if comma:
        decimal = text.replace(",", ".")
        # 1,234 is 1.234 in French and 1234 in English: look for both.
        if _grouped(text, ","):
            return [decimal, text.replace(",", "")]
        return [decimal]
    if dot > 1:
        return [text.replace(".", "")] if _grouped(text, ".") else []
    return [text]


def displayed_numbers(query: str) -> list[tuple[float, float]]:
    """Inclusive bounds of the saved numbers a numeric query stands for.

    Empty when the query is not a number. A percentage also stands for its
    fraction, since Excel saves 34.3 % as 0.343.
    """
    text = query.strip()
    percent = text.endswith("%")
    if percent:
        text = text[:-1].strip()
    text = _GROUP.sub("", text)
    ranges: list[tuple[float, float]] = []
    for spelling in _spellings(text):
        found = _range(spelling)
        if found is None:
            continue
        for scale in (1, 100) if percent else (1,):
            item = (found[0] / scale, found[1] / scale)
            if item not in ranges:
                ranges.append(item)
    if ranges or percent:
        return ranges
    try:
        # Other spellings Python reads, such as 1e6, keep an exact match.
        value = float(query.strip())
    except ValueError:
        return []
    if not math.isfinite(value):
        return []
    margin = abs(value) * 1e-12
    return [(value - margin, value + margin)]


def cell_address(query: str) -> tuple[str | None, int, int, str] | None:
    """``(sheet or None, row, column, A1)`` when the query is one cell address."""
    match = _ADDRESS.fullmatch(query.strip())
    if not match:
        return None
    quoted, bare, letters, digits = match.groups()
    column = 0
    for letter in letters.upper():
        column = column * 26 + ord(letter) - 64
    row = int(digits)
    if row > MAX_ROW or column > MAX_COLUMN:
        return None
    sheet = quoted.replace("''", "'") if quoted else bare
    return sheet, row, column, f"{letters.upper()}{row}"
