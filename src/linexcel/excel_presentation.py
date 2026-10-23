"""Bounded, static worksheet presentation, separate from source data semantics.

Only supported number formats are interpreted. This is a browser preview, not
an Excel renderer: conditional formatting, drawings and print layout are omitted.
All values and formula provenance still come from the imported source index.
"""

from __future__ import annotations

import colorsys
import datetime as dt
import json
import math
import re
import zipfile
import zlib
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation, localcontext
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
    from_excel,
)

from linexcel.formula_source import _worksheet_parts
from linexcel.lazy_store import WorkbookIndex, _unpack

MAX_ROWS = 100
MAX_COLUMNS = 40
MAX_METADATA_BYTES = 4 * 1024 * 1024
MAX_WORKSHEET_BYTES = 64 * 1024 * 1024
MAX_ELEMENTS = 1_000_000
MAX_STYLES = 8192
MAX_MERGES = 2000
MAX_FORMAT_LENGTH = 1024
MAX_PRESENTATION_BYTES = 2 * 1024 * 1024


class PresentationLimitError(RuntimeError):
    """Source presentation exceeds an explicit inspection budget."""


def _bounded_node(packed, remaining):
    """Enforce the window budget before inflating or retaining source records."""
    try:
        decoder = zlib.decompressobj()
        decoded = decoder.decompress(packed, remaining + 1)
    except zlib.error as exc:
        raise ValueError("Cellule indexée illisible") from exc
    if len(decoded) > remaining or decoder.unconsumed_tail:
        raise PresentationLimitError("Zone de présentation trop volumineuse")
    if not decoder.eof:
        raise ValueError("Cellule indexée incomplète")
    # The existing decoder expands short keys and restores saved value aliases.
    # Its second inflation is safe after the complete stream passed this bound.
    return _unpack(packed), len(decoded)


def _xml(archive, part):
    if archive.getinfo(part).file_size > MAX_METADATA_BYTES:
        raise PresentationLimitError("Métadonnées de présentation trop volumineuses")
    data = archive.read(part)
    if b"\x00" in data or b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ValueError("Déclaration XML non prise en charge")
    return ET.fromstring(data)


class _SafeXML:
    """Reject entity declarations even when split between streaming reads."""

    def __init__(self, source):
        self.source, self.tail = source, b""

    def read(self, size=-1):
        data = self.source.read(size)
        combined = self.tail + data.upper()
        if b"\x00" in combined or b"<!DOCTYPE" in combined or b"<!ENTITY" in combined:
            raise ValueError("Déclaration XML non prise en charge")
        self.tail = combined[-16:]
        return data


def _number(value, default, minimum, maximum):
    try:
        number = float(value)
        return max(minimum, min(maximum, number)) if math.isfinite(number) else default
    except (ValueError, TypeError):
        return default


def _integer(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def _enabled(element, key="val", default=True):
    return element is not None and element.get(key, "1" if default else "0") not in {
        "0",
        "false",
        "False",
    }


def _theme(archive):
    if "xl/theme/theme1.xml" not in archive.namelist():
        return []
    scheme = _xml(archive, "xl/theme/theme1.xml").find("{*}themeElements/{*}clrScheme")
    if scheme is None:
        return []
    # OOXML theme indices are lt1, dk1, lt2, dk2, then accents/hyperlinks.
    names = (
        "lt1",
        "dk1",
        "lt2",
        "dk2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    )
    result = []
    for name in names:
        item = scheme.find("{*}" + name)
        color = next(iter(item), None) if item is not None else None
        result.append(
            color.get("lastClr", color.get("val", "")) if color is not None else ""
        )
    return result


def _color(element, theme):
    if element is None:
        return None
    value = element.get("rgb")
    if value is None and "theme" in element.attrib:
        index = _integer(element.get("theme"), -1)
        value = theme[index] if 0 <= index < len(theme) else None
    if value is None and "indexed" in element.attrib:
        index = _integer(element.get("indexed"), -1)
        value = COLOR_INDEX[index] if 0 <= index < len(COLOR_INDEX) else None
    if not isinstance(value, str) or not re.fullmatch(
        r"(?:[0-9A-Fa-f]{8}|[0-9A-Fa-f]{6})", value
    ):
        return None
    value = value[-6:]
    tint = _number(element.get("tint"), 0, -1, 1)
    if tint:
        rgb = [int(value[offset : offset + 2], 16) / 255 for offset in (0, 2, 4)]
        hue, luminance, saturation = colorsys.rgb_to_hls(*rgb)
        luminance = (
            luminance * (1 + tint) if tint < 0 else luminance * (1 - tint) + tint
        )
        rgb = colorsys.hls_to_rgb(hue, luminance, saturation)
        value = "".join(f"{round(channel * 255):02X}" for channel in rgb)
    return "#" + value.upper()


def _styles(archive):
    if "xl/styles.xml" not in archive.namelist():
        return {0: {"css": {}, "numberFormat": "General"}}
    root, theme = _xml(archive, "xl/styles.xml"), _theme(archive)
    collections = {}
    for name in ("fonts", "fills", "borders", "cellXfs"):
        parent = root.find("{*}" + name)
        collections[name] = list(parent) if parent is not None else []
        if len(collections[name]) > MAX_STYLES:
            raise PresentationLimitError("Trop de styles de présentation")
    formats = dict(BUILTIN_FORMATS)
    for item in root.findall("{*}numFmts/{*}numFmt"):
        code = item.get("formatCode", "General")
        if len(code) > MAX_FORMAT_LENGTH:
            raise PresentationLimitError("Format numérique trop long pour cet aperçu")
        formats[_integer(item.get("numFmtId"))] = code
    result = {}
    for index, xf in enumerate(collections["cellXfs"]):
        css = {}
        font_id, fill_id, border_id = (
            _integer(xf.get(name)) for name in ("fontId", "fillId", "borderId")
        )
        if 0 <= font_id < len(collections["fonts"]):
            font = collections["fonts"][font_id]
            if _enabled(font.find("{*}b")):
                css["fontWeight"] = "bold"
            if _enabled(font.find("{*}i")):
                css["fontStyle"] = "italic"
            decorations = []
            if _enabled(font.find("{*}u")):
                decorations.append("underline")
            if _enabled(font.find("{*}strike")):
                decorations.append("line-through")
            if decorations:
                css["textDecoration"] = " ".join(decorations)
            size = font.find("{*}sz")
            if size is not None:
                css["fontSize"] = f"{_number(size.get('val'), 11, 6, 72):g}pt"
            name = font.find("{*}name")
            if name is not None and re.fullmatch(
                r"[A-Za-z0-9 _-]{1,80}", name.get("val", "")
            ):
                css["fontFamily"] = name.get("val")
            color = _color(font.find("{*}color"), theme)
            if color:
                css["color"] = color
        if 0 <= fill_id < len(collections["fills"]):
            pattern = collections["fills"][fill_id].find("{*}patternFill")
            if pattern is not None and pattern.get("patternType") == "solid":
                color = _color(pattern.find("{*}fgColor"), theme)
                if color:
                    css["backgroundColor"] = color
        if 0 <= border_id < len(collections["borders"]):
            border = collections["borders"][border_id]
            for side in ("left", "right", "top", "bottom"):
                edge = border.find("{*}" + side)
                if edge is None or not edge.get("style"):
                    continue
                kind = edge.get("style")
                width = (
                    "3px"
                    if kind == "thick"
                    else "2px"
                    if kind.startswith("medium")
                    else "1px"
                )
                line = (
                    "double"
                    if kind == "double"
                    else "dotted"
                    if kind == "dotted"
                    else "dashed"
                    if "dash" in kind.lower()
                    else "solid"
                )
                color = _color(edge.find("{*}color"), theme) or "#000000"
                css["border" + side.title()] = f"{width} {line} {color}"
        alignment = xf.find("{*}alignment")
        if alignment is not None:
            horizontal = alignment.get("horizontal")
            if horizontal in {"left", "center", "right", "justify"}:
                css["textAlign"] = horizontal
            vertical = alignment.get("vertical")
            if vertical in {"top", "center", "bottom"}:
                css["verticalAlign"] = "middle" if vertical == "center" else vertical
            if _enabled(alignment, "wrapText", False):
                css["whiteSpace"] = "pre-wrap"
        result[index] = {
            "css": css,
            "numberFormat": formats.get(_integer(xf.get("numFmtId")), "General"),
        }
    return result or {0: {"css": {}, "numberFormat": "General"}}


def format_display(value, number_format, *, raw=None, epoch1904=False):
    """Render a deliberately small Excel format subset; label every fallback."""
    plain = (
        ""
        if value is None
        else "TRUE"
        if value is True
        else "FALSE"
        if value is False
        else str(value)
    )
    if value is None or isinstance(value, bool):
        return plain, "raw"
    if number_format in {None, "", "General", "@"}:
        return plain, "raw"
    if isinstance(value, str) and not is_date_format(number_format):
        return plain, "raw"
    try:
        if is_date_format(number_format):
            # Numeric, locale-independent dates and times only. Month/day names,
            # elapsed durations and locale annotations require a real renderer.
            code = number_format.lower()
            if re.search(r"[^ymdhs :./-]", code.replace("am/pm", "")) or re.search(
                r"d{3,}|m{3,}|y{5,}|h{3,}|s{3,}", code
            ):
                return plain, "unsupported"
            if raw is None:
                return plain, "unsupported"
            serial = float(raw)
            if not epoch1904 and (serial < 0 or 60 <= serial < 61):
                return plain, "unsupported"
            epoch = CALENDAR_MAC_1904 if epoch1904 else CALENDAR_WINDOWS_1900
            date = from_excel(serial, epoch)
            if isinstance(date, dt.time) and re.search("[yd]", code):
                if not epoch1904:
                    return plain, "unsupported"
                date = epoch + dt.timedelta(days=serial)
            twelve_hour = "am/pm" in code
            tokens = list(re.finditer(r"am/pm|y+|m+|d+|h+|s+", code))
            result, previous = [], 0
            for position, token in enumerate(tokens):
                result.append(number_format[previous : token.start()])
                key, width = token[0][0], len(token[0])
                if token[0] == "am/pm":
                    part = "AM" if date.hour < 12 else "PM"
                elif key == "y":
                    if width not in {2, 4}:
                        return plain, "unsupported"
                    part = str(date.year) if width > 2 else f"{date.year % 100:02d}"
                else:
                    minute = key == "m" and (
                        (position > 0 and tokens[position - 1][0].startswith("h"))
                        or (
                            position + 1 < len(tokens)
                            and tokens[position + 1][0].startswith("s")
                        )
                    )
                    attribute = (
                        "minute"
                        if minute
                        else {"m": "month", "d": "day", "h": "hour", "s": "second"}[key]
                    )
                    number = getattr(date, attribute)
                    if key == "h" and twelve_hour:
                        number = number % 12 or 12
                    part = str(number).zfill(2 if width > 1 else 1)
                result.append(part)
                previous = token.end()
            result.append(number_format[previous:])
            return "".join(result), "formatted"
        if not isinstance(value, (int, float)):
            return plain, "unsupported"
        # Split outside quoted literals; conditions and locale directives fail
        # closed. Color annotations do not change the textual result.
        sections = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', number_format)
        section = (
            sections[2]
            if value == 0 and len(sections) > 2
            else sections[1]
            if value < 0 and len(sections) > 1
            else sections[0]
        )
        section = re.sub(
            r"\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]",
            "",
            section,
            flags=re.I,
        )
        if "[" in section or "]" in section or "_" in section or "*" in section:
            return plain, "unsupported"
        literals = {}

        def literal(match):
            key = chr(0xE000 + len(literals))
            literals[key] = match[0][1:-1] if match[0].startswith('"') else match[0][1:]
            return key

        section = re.sub(r'"[^\"]*"|\\.', literal, section)
        match = re.fullmatch(r"([^0#?,%]*)(0|#,##0)(\.0+)?(%?)([^0#?,%]*)", section)
        if not match:
            return plain, "unsupported"
        prefix, pattern, fraction, percent, suffix = match.groups()
        if value < 0 and len(sections) == 1 and prefix:
            return plain, "unsupported"
        if re.search(r"[A-Za-z@]", prefix + suffix):
            return plain, "unsupported"
        decimals = len(fraction) - 1 if fraction else 0
        if decimals > 15:
            return plain, "unsupported"
        with localcontext() as context:
            context.prec = 340
            absolute = Decimal(str(value))
            absolute = abs(absolute) if value < 0 and len(sections) > 1 else absolute
            absolute *= 100 if percent else 1
            absolute = absolute.quantize(
                Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP
            )
            formatted = format(absolute, f"{',' if ',' in pattern else ''}.{decimals}f")
        result = prefix + formatted + percent + suffix
        for key, text in literals.items():
            result = result.replace(key, text)
        return result, "formatted"
    except (ValueError, TypeError, OverflowError, AttributeError, InvalidOperation):
        return plain, "unsupported"


def sheet_presentation(
    workbook: Path, index: Path, sheet: str, *, row=1, column=1, rows=60, columns=20
):
    """Read one bounded worksheet window without evaluating any formula."""
    if not 1 <= row <= 1048576 or not 1 <= column <= 16384:
        raise ValueError("Coordonnées Excel hors limites")
    rows, columns = max(1, min(MAX_ROWS, rows)), max(1, min(MAX_COLUMNS, columns))
    last_row, last_column = (
        min(1048576, row + rows - 1),
        min(16384, column + columns - 1),
    )
    with WorkbookIndex(index) as source:
        if sheet not in source.meta.get("sheets", []):
            raise KeyError(sheet)
        sha = source.meta.get("sha256")
        nodes, node_bytes = {}, 0
        for size, packed in source.db.execute(
            "SELECT length(data), CASE WHEN length(data)<=? THEN data END "
            "FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
            "AND col BETWEEN ? AND ? AND kind IN ('cell','input')",
            (MAX_PRESENTATION_BYTES + 65536, sheet, row, last_row, column, last_column),
        ):
            if packed is None:
                raise PresentationLimitError("Cellule indexée trop volumineuse")
            node, decoded_bytes = _bounded_node(
                packed, MAX_PRESENTATION_BYTES - node_bytes
            )
            node_bytes += decoded_bytes
            if (node.get("sourceValueProvenance") or {}).get("workbookSha256") == sha:
                nodes[(node["row"], node["column"])] = node
    with zipfile.ZipFile(workbook) as archive:
        # Bound metadata before reusing the workbook relationship reader.
        _xml(archive, "xl/workbook.xml")
        _xml(archive, "xl/_rels/workbook.xml.rels")
        part = next(
            (part for name, part in _worksheet_parts(archive) if name == sheet), None
        )
        if part is None:
            raise ValueError("Source OOXML de la feuille introuvable")
        if archive.getinfo(part).file_size > MAX_WORKSHEET_BYTES:
            raise PresentationLimitError(
                "Feuille trop volumineuse pour cet aperçu statique"
            )
        styles = _styles(archive)
        workbook_properties = _xml(archive, "xl/workbook.xml").find("{*}workbookPr")
        epoch1904 = _enabled(workbook_properties, "date1904", False)
        cells, row_dimensions, column_dimensions, merges = {}, {}, {}, []
        extent = {"rows": 0, "columns": 0}
        defaults = {"rowHeight": 20, "columnWidth": 64}
        freeze_pane = None
        notices = [
            "Aperçu statique des valeurs sauvegardées ; aucun recalcul.",
            "Formats numériques courants seulement ; "
            "les formats complexes restent bruts. "
            "Mises en forme conditionnelles, dessins, graphiques, rotation "
            "couleurs des formats numériques et mise en page d’impression "
            "non reproduits.",
            "Largeurs, hauteurs et teintes de thème approximatives "
            "dans le navigateur ; "
            "volets figés extraits mais non appliqués.",
        ]
        with archive.open(part) as handle:
            stack, count, current_row, previous_column = [], 0, 0, 0
            for event, element in ET.iterparse(
                _SafeXML(handle), events=("start", "end")
            ):
                if event == "start":
                    count += 1
                    if count > MAX_ELEMENTS:
                        raise PresentationLimitError(
                            "Feuille trop complexe pour cet aperçu statique"
                        )
                    stack.append(element)
                    if element.tag.rsplit("}", 1)[-1] == "row":
                        current_row = _integer(element.get("r"), current_row + 1)
                        previous_column = 0
                    continue
                tag = element.tag.rsplit("}", 1)[-1]
                if tag == "sheetFormatPr":
                    defaults["rowHeight"] = (
                        _number(element.get("defaultRowHeight"), 15, 1, 300) * 4 / 3
                    )
                    defaults["columnWidth"] = (
                        _number(element.get("defaultColWidth"), 8.43, 0, 255) * 7 + 5
                    )
                elif tag == "pane":
                    freeze_pane = {
                        key: element.get(key)
                        for key in ("state", "topLeftCell", "xSplit", "ySplit")
                        if element.get(key) is not None
                    }
                elif tag == "col":
                    first, last = (
                        _integer(element.get("min")),
                        _integer(element.get("max")),
                    )
                    for col in range(max(column, first), min(last_column, last) + 1):
                        column_dimensions[str(col)] = {
                            "width": _number(element.get("width"), 8.43, 0, 255) * 7
                            + 5,
                            "hidden": _enabled(element, "hidden", False),
                            **(
                                {"styleId": _integer(element.get("style"))}
                                if element.get("style") is not None
                                else {}
                            ),
                        }
                elif tag == "c":
                    try:
                        address = element.get("r")
                        if address is None:
                            cell_row, cell_column = current_row, previous_column + 1
                        else:
                            cell_row, cell_column = coordinate_to_tuple(address)
                    except (ValueError, TypeError):
                        cell_row, cell_column = 0, 0
                    previous_column = cell_column
                    if 1 <= cell_row <= 1048576 and 1 <= cell_column <= 16384:
                        extent["rows"] = max(extent["rows"], cell_row)
                        extent["columns"] = max(extent["columns"], cell_column)
                    if (
                        row <= cell_row <= last_row
                        and column <= cell_column <= last_column
                    ):
                        parent = stack[-2] if len(stack) > 1 else None
                        fallback = parent.get("s") if parent is not None else None
                        if fallback is None:
                            fallback = column_dimensions.get(str(cell_column), {}).get(
                                "styleId", 0
                            )
                        cells[(cell_row, cell_column)] = _integer(
                            element.get("s"), _integer(fallback)
                        )
                elif tag == "row":
                    cell_row = current_row
                    if row <= cell_row <= last_row:
                        row_dimensions[str(cell_row)] = {
                            "height": _number(
                                element.get("ht"), defaults["rowHeight"] * 3 / 4, 0, 300
                            )
                            * 4
                            / 3,
                            "hidden": _enabled(element, "hidden", False),
                            **(
                                {"styleId": _integer(element.get("s"))}
                                if element.get("s") is not None
                                else {}
                            ),
                        }
                elif tag == "mergeCell":
                    try:
                        c1, r1, c2, r2 = range_boundaries(element.get("ref", ""))
                        if (
                            c1 is not None
                            and r1 is not None
                            and c2 is not None
                            and r2 is not None
                            and r1 <= last_row
                            and r2 >= row
                            and c1 <= last_column
                            and c2 >= column
                        ):
                            if len(merges) >= MAX_MERGES:
                                raise PresentationLimitError(
                                    "Trop de cellules fusionnées dans cet aperçu"
                                )
                            merges.append(
                                {
                                    "range": element.get("ref"),
                                    "row": r1,
                                    "column": c1,
                                    "lastRow": r2,
                                    "lastColumn": c2,
                                }
                            )
                            extent["rows"] = max(extent["rows"], r2)
                            extent["columns"] = max(extent["columns"], c2)
                    except (ValueError, TypeError):
                        continue
                # Retain a complete cell until its end, but never retain rows.
                if tag in {"c", "row", "col", "mergeCell"}:
                    if len(stack) > 1:
                        stack[-2].remove(element)
                    element.clear()
                stack.pop()
        # Row/column styles also apply to implicit blanks. Materialize only
        # styled blanks inside the bounded window, never an entire worksheet.
        for r in range(row, last_row + 1):
            row_style = row_dimensions.get(str(r), {}).get("styleId")
            for c in range(column, last_column + 1):
                style_id = (
                    row_style
                    if row_style is not None
                    else (column_dimensions.get(str(c), {}).get("styleId", 0))
                )
                if style_id:
                    cells.setdefault((r, c), style_id)
        records, used_styles = [], {0}
        record_bytes = 0
        for dimensions in (row_dimensions, column_dimensions):
            used_styles.update(
                dimension["styleId"]
                for dimension in dimensions.values()
                if "styleId" in dimension
            )
        for r, c in sorted(cells.keys() | nodes.keys()):
            style_id = cells.get((r, c), 0)
            node = nodes.get((r, c), {})
            style = styles.get(
                style_id, styles.get(0, {"css": {}, "numberFormat": "General"})
            )
            value, provenance = (
                node.get("value"),
                node.get("sourceValueProvenance") or {},
            )
            display, status = format_display(
                value,
                style["numberFormat"],
                raw=provenance.get("rawXmlValue"),
                epoch1904=epoch1904,
            )
            used_styles.add(style_id)
            record = {
                "row": r,
                "column": c,
                "address": f"{get_column_letter(c)}{r}",
                "value": value,
                "display": display,
                "formatStatus": status,
                "styleId": style_id,
                "numberFormat": style["numberFormat"],
                "formula": node.get("sourceFormula") or node.get("canonicalFormula"),
                "sourceFormula": node.get("sourceFormula"),
                "canonicalFormula": node.get("canonicalFormula"),
                "formulaRole": "source"
                if node.get("sourceFormula")
                else ("canonical" if node.get("canonicalFormula") else "unavailable"),
                "valueStatus": provenance.get("status", "blank"),
            }
            record_bytes += len(json.dumps(record, ensure_ascii=False).encode("utf-8"))
            if record_bytes > MAX_PRESENTATION_BYTES:
                raise PresentationLimitError("Zone de présentation trop volumineuse")
            records.append(record)
        result = {
            "sheet": sheet,
            "sourceSha256": sha,
            "cells": records,
            "styles": {str(key): styles[key] for key in used_styles if key in styles},
            "rowDimensions": row_dimensions,
            "columnDimensions": column_dimensions,
            "merges": merges,
            "extent": extent,
            "defaults": defaults,
            "freezePane": freeze_pane,
            "window": {
                "row": row,
                "column": column,
                "rows": last_row - row + 1,
                "columns": last_column - column + 1,
            },
            "limits": {
                "maxRows": MAX_ROWS,
                "maxColumns": MAX_COLUMNS,
                "worksheetBytes": MAX_WORKSHEET_BYTES,
                "responseBytes": MAX_PRESENTATION_BYTES,
            },
            "notices": notices,
        }
        if (
            len(json.dumps(result, ensure_ascii=False).encode("utf-8"))
            > MAX_PRESENTATION_BYTES
        ):
            raise PresentationLimitError("Zone de présentation trop volumineuse")
        return result
