"""Read a search query the way a reader types it: a displayed number or an address.

Excel shows 967.5 for a saved 967.538 and 6600 for 6600.000000000001. A reader
searches what the sheet shows, so a number matches every saved value that
rounds to it at the number of decimals typed.
"""

from __future__ import annotations

import math
import re
import sqlite3
import time

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


CELL_KINDS = "kind IN ('cell','input','array_result')"


def cells_at(revisions, index_of, address, fold, deadline, limit, caps):
    """Cells at one address in each scoped sheet, read from the source indexes.

    Returns ``(revision item, sheet, A1, node id, node)`` tuples. ``index_of``
    opens the index of a revision item, ``fold`` is the search's own text
    normalization and ``caps`` collects the limits reached.
    """
    wanted, row, column, a1 = address
    folded = fold(wanted) if wanted is not None else None
    found = []
    for item in revisions:
        if not item["revision"] or item["rebuildRequired"]:
            continue
        index = index_of(item)
        for sheet in index.meta.get("sheets", []):
            if folded is not None and fold(sheet) != folded:
                continue
            if time.monotonic() >= deadline:
                caps.add("query_ms")
                return found
            for nid in index.range_members(sheet, (row, column, row, column)):
                found.append((item, sheet, a1, nid, index[nid]))
        if len(found) > limit:
            caps.add("hits")
            break
    return found


def numbers_in(revisions, index_of, ranges, seen, deadline, limit, caps):
    """Saved numbers inside ``ranges`` that the labelled index did not return.

    The catalog indexes numbers that have a label. In a wide table most values
    sit too far from their row label to get one, and a reader still searches the
    number the sheet shows. One pass over each source index finds them; the
    time budget of the query bounds it.
    """
    bounds = [bound for item in ranges for bound in item]
    clause = " OR ".join("CAST(search_value AS REAL) BETWEEN ? AND ?" for _ in ranges)
    found = []
    for item in revisions:
        if not item["revision"] or item["rebuildRequired"]:
            continue
        if len(found) > limit:
            caps.add("hits")
            break
        index = index_of(item)
        index.db.set_progress_handler(lambda: int(time.monotonic() >= deadline), 1000)
        try:
            # Digits and number punctuation only: a text such as "12 rue" casts
            # to 12 and must not be read as a number.
            cursor = index.db.execute(
                f"SELECT id FROM nodes WHERE {CELL_KINDS} "
                "AND search_value GLOB '*[0-9]*' "
                "AND search_value NOT GLOB '*[^0-9.eE+-]*' "
                f"AND ({clause}) ORDER BY seq LIMIT ?",
                (*bounds, limit + 1),
            )
            for (nid,) in cursor.fetchall():
                if (item["projectId"], nid) in seen:
                    continue
                node = index[nid]
                value = node.get("value")
                if type(value) not in (int, float) or not any(
                    low <= value <= high for low, high in ranges
                ):
                    continue
                address = node.get("cell") or node.get("reference")
                found.append((item, node.get("sheet"), address, nid, node))
        except sqlite3.OperationalError as error:
            if "interrupt" not in str(error):
                raise
            caps.add("query_ms")
        finally:
            index.db.set_progress_handler(None, 0)
    return found
