"""Bounded source inventory after OOM; never recalculate a truncated workbook."""

from __future__ import annotations

import io
import math
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl.formula.translate import Translator
from openpyxl.utils.escape import unescape

from linexcel.loader import _parse_sheet_targets
from linexcel.progress import Reporter

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
MAX_TEXT = 32_767


def _decode_text(value: str) -> str | None:
    # Decode once: _x005F_x0041_ is literal _x0041_, not the letter A.
    # Escaped UTF-16 surrogate pairs must become serializable Unicode text.
    try:
        return unescape(value).encode("utf-16-le", "surrogatepass").decode("utf-16-le")
    except UnicodeDecodeError:
        return None


def _visible_text(element) -> str | None:
    # rPh contains pronunciation guides, not the visible stored cell value.
    if element is None:
        return ""
    return _decode_text(
        "".join(
            (child.text or "")
            if child.tag == NS + "t"
            else child.findtext(NS + "t", "")
            for child in element
            if child.tag in {NS + "t", NS + "r"}
        )
    )


def _cells(stream):
    """Read stored cells, not the dense rectangle implied by dimensions."""
    parents = []
    for event, element in ET.iterparse(stream, events=("start", "end")):
        if event == "start":
            parents.append(element)
            continue
        if element.tag == NS + "c":
            yield element
        if element.tag in {NS + "c", NS + "row"}:
            if len(parents) > 1:
                parents[-2].remove(element)
            element.clear()
        parents.pop()


def _value(cell) -> tuple[Any, int | None]:
    kind = cell.get("t", "n")
    raw = cell.findtext(NS + "v")
    if kind == "inlineStr":
        raw = _visible_text(cell.find(NS + "is"))
    if raw is None or len(raw) > MAX_TEXT:
        return None, None
    if kind == "s":
        try:
            index = int(raw)
            return None, index if index >= 0 else None
        except ValueError:
            return None, None
    if kind in {"str", "inlineStr", "e", "d"}:
        return (_decode_text(raw) if kind == "str" else raw), None
    if kind == "b":
        return (raw == "1" if raw in {"0", "1"} else None), None
    try:
        number = float(raw)
        return (number if math.isfinite(number) else None), None
    except ValueError:
        return None, None


def source_inventory(
    data: bytes, kwargs: dict, cell_budget: int, reporter: Reporter
) -> dict:
    """Return a sample of formulas/file values, explicitly without lineage.

    Budgets count stored cell elements, including styled blanks. Limits are
    divided across worksheets before reading any, so a dense first sheet
    cannot consume every slot. No native parser/evaluator is called here.
    """
    nodes: list[dict[str, Any]] = []
    sheets = []
    stats: list[dict[str, Any]] = []
    pending_strings: dict[int, list[dict]] = {}
    omissions = [{"phase": "recovery", "reason": "dependencies_not_inspected"}]
    warnings = [
        "Partial source-only recovery after memory exhaustion: formulas and "
        "stored values only. No recalculation, dependency tracing, decomposition, "
        "VBA or Power Query analysis was performed. Missing stored values remain "
        "unknown. Numeric values use the file's raw encoding (dates may be Excel "
        "serial numbers); formatting is not reconstructed."
    ]
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        paths = _parse_sheet_targets(
            archive.read("xl/workbook.xml").decode("utf-8"),
            archive.read("xl/_rels/workbook.xml.rels").decode("utf-8"),
        )
        # Chartsheets are not cell worksheets.
        paths = {s: p for s, p in paths.items() if "/worksheets/" in p}
        sheets = list(paths)
        per_sheet, extra = divmod(cell_budget, max(1, len(paths)))
        for index, (sheet, part) in enumerate(paths.items()):
            quota = per_sheet + (index < extra)
            for key, default in (
                ("max_cells_per_sheet", quota),
                ("max_nodes_per_sheet", 400),
            ):
                limit = kwargs.get(key)
                quota = min(quota, default if limit is None else limit)
            reporter.operation("source recovery", f"reading stored cells on {sheet}")
            scanned = formulas = 0
            limited = False
            shared: dict[str, tuple[str, str]] = {}
            with archive.open(part) as stream:
                for cell in _cells(stream):
                    if scanned >= quota:
                        limited = True
                        break
                    scanned += 1
                    address = cell.get("r")
                    if not address:
                        continue
                    f = cell.find(NS + "f")
                    value, string_index = _value(cell)
                    if f is None and value is None and string_index is None:
                        continue
                    node: dict[str, Any] = {
                        "id": f"{'c' if f is not None else 'i'}:{sheet}!{address}",
                        "kind": "cell" if f is not None else "input",
                        "sheet": sheet,
                        "addr": address,
                        "label": f"{sheet}!{address}",
                        "value": value,
                        "valueSource": "file" if value is not None else None,
                        "steps": None,
                        "recoveryOnly": True,
                    }
                    if f is not None:
                        formulas += 1
                        text = f.text or ""
                        if text and len(text) <= MAX_TEXT:
                            node["formula"] = "=" + text
                            if f.get("t") == "shared":
                                shared[f.get("si", "")] = (address, "=" + text)
                        elif f.get("t") == "shared" and f.get("si", "") in shared:
                            origin, formula = shared[f.get("si", "")]
                            try:
                                node["formula"] = Translator(
                                    formula, origin=origin
                                ).translate_formula(address)
                            except Exception:
                                node["formulaUnavailable"] = True
                        else:
                            node["formulaUnavailable"] = True
                    if string_index is not None:
                        pending_strings.setdefault(string_index, []).append(node)
                    nodes.append(node)
            stats.append(
                {
                    "name": sheet,
                    "recoveredCells": scanned,
                    "recoveredFormulas": formulas,
                    "cellLimit": quota,
                    "truncated": limited,
                }
            )
            if limited:
                omissions.append(
                    {
                        "phase": "recovery",
                        "sheet": sheet,
                        "reason": "stored_cell_limit",
                        "limit": quota,
                        "omittedCellCount": None,
                    }
                )
        if pending_strings and "xl/sharedStrings.xml" in archive.namelist():
            with archive.open("xl/sharedStrings.xml") as stream:
                iterator = ET.iterparse(stream, events=("start", "end"))
                _, root = next(iterator)
                index = 0
                for event, element in iterator:
                    if event != "end" or element.tag != NS + "si":
                        continue
                    if index in pending_strings:
                        text = _visible_text(element)
                        if text is not None and len(text) <= MAX_TEXT:
                            for node in pending_strings[index]:
                                node["value"] = text
                                node["valueSource"] = "file"
                        del pending_strings[index]
                    index += 1
                    root.remove(element)
                    element.clear()
                    if not pending_strings:
                        break
    warnings.append(
        f"Recovery inspected {sum(s['recoveredCells'] for s in stats)} stored "
        f"cell(s) across {len(sheets)} worksheet(s), with a total budget of "
        f"{cell_budget}. Per-sheet limits and omissions are recorded in coverage."
    )
    return {
        "meta": {
            "filename": kwargs.get("filename", "workbook.xlsx"),
            "warnings": warnings,
            "engine": "source-only recovery (no recalculation)",
            "stats": {
                "totalNodes": len(nodes),
                "totalEdges": 0,
                "totalFormulas": None,
                "sheets": stats,
            },
            "analysisCoverage": {
                "requested": "workbook",
                "recoveryMode": "source_only",
                "dependencyCompleteness": "not_inspected",
                "extractedFormulaCells": sum(s["recoveredFormulas"] for s in stats),
                "omissions": omissions,
            },
        },
        "nodes": nodes,
        "edges": [],
        "sheets": sheets,
    }
