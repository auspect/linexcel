"""Bounded, source-based resolution of local Excel table references.

The resolver never infers a table from formatted cells. Its rectangles describe
the declared OOXML table, independently of saved values or formula results.
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.etree import ElementTree as ET

from openpyxl.formula import Tokenizer

from linexcel.refs import Rect, parse_ref, split_sheet_prefix

MAX_TABLE_PART_BYTES = 4 * 1024 * 1024


def formula_tokens(formula: str):
    """Keep escaped table brackets literal for openpyxl's bracket scanner.

    Reversible placeholders affect tokenization only; returned tokens retain
    the exact input spelling, including text literals and external references.
    """
    marker = "_linexcel_literal_bracket_"
    while marker in formula:
        marker += "_"
    protected = []
    index = 0
    depth = 0
    while index < len(formula):
        character = formula[index]
        if character == '"' or (character == "'" and not depth):
            quote = character
            start = index
            index += 1
            while index < len(formula):
                if formula[index] == quote:
                    index += 1
                    if index < len(formula) and formula[index] == quote:
                        index += 1
                        continue
                    break
                index += 1
            protected.append(formula[start:index])
            continue
        if depth and character == "'" and index + 1 < len(formula):
            following = formula[index + 1]
            if following in "[]":
                protected.append(marker + ("open" if following == "[" else "close"))
                index += 2
                continue
            if following in "#'@":
                protected.append(formula[index : index + 2])
                index += 2
                continue
        if character == "[":
            depth += 1
        elif character == "]":
            depth = max(0, depth - 1)
        elif (
            character == "#"
            and not depth
            and index
            and formula[index - 1] in "0123456789]"
        ):
            protected.append(marker + "spill")
            index += 1
            continue
        protected.append(character)
        index += 1
    tokens = Tokenizer("".join(protected)).items
    for token in tokens:
        token.value = (
            token.value.replace(marker + "open", "'[")
            .replace(marker + "close", "']")
            .replace(marker + "spill", "#")
        )
    return tokens


class TableReferenceError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def read_table(archive, part: str, sheet: str, digest: str) -> dict:
    """Read one bounded table part, rejecting incomplete calculation metadata."""
    with archive.open(part) as source:
        content = source.read(MAX_TABLE_PART_BYTES + 1)
    if len(content) > MAX_TABLE_PART_BYTES:
        raise TableReferenceError(
            "table_metadata_limit", "Table metadata exceeds its byte limit."
        )
    declarations = content.replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in declarations or b"<!ENTITY" in declarations:
        raise TableReferenceError(
            "table_metadata_invalid", "Table entity declarations are unsupported."
        )
    root = ET.fromstring(content)
    if root.tag.rsplit("}", 1)[-1] != "table":
        raise TableReferenceError(
            "table_metadata_invalid", "Expected an OOXML table definition."
        )
    name = root.get("displayName") or root.get("name")
    reference = root.get("ref", "")
    rect = parse_ref(reference)
    if not name or rect is None or rect.sheet is not None:
        raise TableReferenceError(
            "table_metadata_invalid", "Table name or rectangle is invalid."
        )
    header_rows = int(root.get("headerRowCount", "1"))
    total_rows = int(root.get("totalsRowCount", "0"))
    columns = root.findall("{*}tableColumns/{*}tableColumn")
    headers = [column.get("name") for column in columns]
    if (
        header_rows not in {0, 1}
        or total_rows not in {0, 1}
        or header_rows + total_rows > rect.r2 - rect.r1 + 1
        or len(headers) != rect.c2 - rect.c1 + 1
        or any(not isinstance(header, str) or not header for header in headers)
        or len({str(header).casefold() for header in headers}) != len(headers)
    ):
        raise TableReferenceError(
            "table_metadata_invalid",
            "Table columns, headers or totals are inconsistent.",
        )
    return {
        "name": name,
        "sheet": sheet,
        "bounds": [rect.r1, rect.c1, rect.r2, rect.c2],
        "headers": headers,
        "headerRows": header_rows,
        "totalRows": total_rows,
        "activeFilters": bool(root.findall(".//{*}filterColumn")),
        "sourcePart": part,
        "sourceSha256": digest,
    }


def _unescape(text: str) -> str:
    result = []
    index = 0
    while index < len(text):
        if text[index] == "'" and index + 1 < len(text) and text[index + 1] in "[]#'@":
            index += 1
        result.append(text[index])
        index += 1
    return "".join(result)


def _bracket(text: str, start: int = 0, depth: int = 0):
    if depth > 3 or start >= len(text) or text[start] != "[":
        raise TableReferenceError(
            "table_reference_syntax", "Unsupported structured-reference syntax."
        )
    index = start + 1
    this_row = index < len(text) and text[index] == "@"
    if this_row:
        index += 1
    parts, separators = [], []
    if index < len(text) and text[index] == "[":
        while True:
            children, child_separators, end = _bracket(text, index, depth + 1)
            parts.extend(children)
            separators.extend(child_separators)
            index = end
            if index < len(text) and text[index] in ",:":
                separators.append(text[index])
                index += 1
            else:
                break
    else:
        begin = index
        while index < len(text):
            if (
                text[index] == "'"
                and index + 1 < len(text)
                and text[index + 1] in "[]#'@"
            ):
                index += 2
                continue
            if text[index] in "[]":
                break
            index += 1
        raw = text[begin:index]
        if raw:
            parts.append((_unescape(raw), raw.startswith("#")))
        elif not this_row:
            raise TableReferenceError(
                "table_reference_syntax", "Empty structured-reference component."
            )
    if index >= len(text) or text[index] != "]":
        raise TableReferenceError(
            "table_reference_syntax", "Unbalanced structured reference."
        )
    if this_row:
        parts.insert(0, ("#This Row", True))
        if len(parts) > 1:
            separators.insert(0, ",")
    return parts, separators, index + 1


@dataclass(frozen=True)
class TableSelection:
    table: dict
    rectangle: Rect


def resolve_table_reference(
    store, text: str, sheet: str | None, row: int | None, column: int | None
) -> TableSelection | None:
    """Resolve a declared table or return None for an ordinary/external token."""
    prefix, body = split_sheet_prefix(text)
    if prefix is not None and "[" in prefix:
        return None
    opening = body.find("[")
    name = body if opening < 0 else body[:opening]
    if opening < 0:
        records = store.db.execute(
            "SELECT data FROM table_definitions WHERE name_key=? LIMIT 2",
            (name.casefold(),),
        ).fetchall()
        if not records:
            return None
    elif name:
        records = store.db.execute(
            "SELECT data FROM table_definitions WHERE name_key=? LIMIT 2",
            (name.casefold(),),
        ).fetchall()
    else:
        records = store.db.execute(
            "SELECT data FROM table_definitions WHERE sheet=? "
            "AND ? BETWEEN r1 AND r2 AND ? BETWEEN c1 AND c2 LIMIT 2",
            (sheet, row, column),
        ).fetchall()
    if len(records) != 1:
        raise TableReferenceError(
            "table_missing" if not records else "table_ambiguous",
            "Unresolved structured reference; no unique declared table: " + text,
        )
    import json

    table = json.loads(records[0][0])
    if prefix is not None and prefix.casefold() != table["sheet"].casefold():
        raise TableReferenceError(
            "table_sheet_mismatch", "Table and sheet qualifier disagree: " + text
        )
    if opening < 0:
        parts, separators = [("#Data", True)], []
    else:
        parts, separators, end = _bracket(body, opening)
        if end != len(body):
            raise TableReferenceError(
                "table_reference_syntax", "Unexpected text after structured reference."
            )
    selectors = [value.casefold() for value, special in parts if special]
    columns = [value for value, special in parts if not special]
    if any(
        value not in {"#data", "#headers", "#totals", "#all", "#this row"}
        for value in selectors
    ):
        raise TableReferenceError(
            "table_selector_unsupported", "Unknown table row selector: " + text
        )
    r1, c1, r2, c2 = table["bounds"]
    first_data, last_data = r1 + table["headerRows"], r2 - table["totalRows"]
    rows = []
    for selector in selectors or ["#data"]:
        if selector == "#all":
            rows.append((r1, r2))
        elif selector == "#data":
            if first_data > last_data:
                raise TableReferenceError(
                    "table_empty", "Empty table data has no source rectangle."
                )
            rows.append((first_data, last_data))
        elif selector == "#headers" and table["headerRows"]:
            rows.append((r1, r1))
        elif selector == "#totals" and table["totalRows"]:
            rows.append((r2, r2))
        elif (
            selector == "#this row"
            and row is not None
            and first_data <= row <= last_data
        ):
            rows.append((row, row))
        else:
            raise TableReferenceError(
                "table_row_unavailable", "Requested table row is unavailable: " + text
            )
    rows.sort()
    lo, hi = rows[0]
    for first, last in rows[1:]:
        if first > hi + 1:
            raise TableReferenceError(
                "table_disjoint_rows",
                "Disjoint table row selections are not supported.",
            )
        hi = max(hi, last)
    if columns:
        names = {name.casefold(): index for index, name in enumerate(table["headers"])}
        try:
            indices = [names[name.casefold()] for name in columns]
        except KeyError as exc:
            raise TableReferenceError(
                "table_column_missing", "Unknown table column: " + str(exc)
            ) from exc
        if len(indices) > 2 or (
            len(indices) == 2
            and (separators.count(":") != 1 or indices[0] > indices[1])
        ):
            raise TableReferenceError(
                "table_columns_unsupported",
                "Unsupported table column selection: " + text,
            )
        c2 = c1 + indices[-1]
        c1 += indices[0]
    elif ":" in separators:
        raise TableReferenceError(
            "table_reference_syntax", "A table column span requires columns."
        )
    return TableSelection(table, Rect(table["sheet"], lo, c1, hi, c2))
