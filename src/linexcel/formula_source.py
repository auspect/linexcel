"""Stored OOXML evidence, kept separate from formulas used for calculation.

``xmlText`` is the decoded character content of ``<f>``, preserving spelling,
whitespace and parentheses. It is not a byte-for-byte copy of XML markup.
Shared slaves have no stored formula text: a translated formula is never
labelled as their original XML source. Legacy graph enrichment retains only
the displayed cells and samples, with a second streamed pass for their masters.
"""

from __future__ import annotations

import hashlib
import io
import posixpath
import re
import zipfile
from collections.abc import Iterable
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter


def cell_source_fields(
    element: ET.Element,
    *,
    workbook_sha256: str,
    worksheet_part: str,
    sheet: str,
    cell: str,
    canonical_formula: str | None,
    canonical_origin: str,
    shared_master: tuple[str, str] | None = None,
    translation_status: str | None = None,
    translated_formula: str | None = None,
    unreadable_value: bool = False,
) -> dict[str, Any]:
    """Describe stored formula/cache evidence without changing semantic data."""
    location = {
        "workbookSha256": workbook_sha256,
        "worksheetPart": worksheet_part,
        "sheet": sheet,
        "cell": cell,
        "xmlCellAddress": element.get("r"),
        "cellAddressOrigin": "stored_xml" if element.get("r") == cell else "inferred",
    }
    formula = element.find("{*}f")
    value = element.find("{*}v")
    storage_type = element.get("t", "n")
    present = value is not None and (value.text is not None or storage_type == "str")
    # An inline string is an input, not a formula result cached in <v>.
    if storage_type == "inlineStr" and formula is None:
        present = element.find("{*}is") is not None
    result: dict[str, Any] = {
        "sourceValueProvenance": {
            **location,
            "role": "formula_cache" if formula is not None else "input",
            "storageType": storage_type,
            "valueElementPresent": value is not None,
            "rawXmlValue": value.text if value is not None else None,
            "status": "unreadable"
            if unreadable_value
            else "present"
            if present
            else "missing",
        }
    }
    if formula is None and canonical_formula is None:
        return result
    xml_text = (formula.text or "") if formula is not None else None
    source_formula = "=" + xml_text if xml_text else None
    provenance: dict[str, Any] = {
        **location,
        "status": "stored_xml"
        if xml_text
        else "stored_empty"
        if formula is not None
        else "unavailable",
        "xmlText": xml_text,
        "xmlTextSemantics": "decoded_xml_character_data",
        "sourceFormulaPrefix": "=" if source_formula is not None else None,
        "sourceFormulaPrefixOrigin": "presentation_added"
        if source_formula is not None
        else "not_available",
        "xmlAttributes": dict(formula.attrib) if formula is not None else {},
        "canonicalOrigin": canonical_origin,
        "canonicalMatchesSource": canonical_formula == source_formula
        if source_formula is not None and canonical_formula is not None
        else None,
    }
    if formula is not None and formula.get("t") == "shared":
        master = bool(xml_text)
        provenance["shared"] = {
            "role": "master" if master else "slave",
            "index": formula.get("si"),
            "masterCell": cell
            if master
            else shared_master[0]
            if shared_master
            else None,
            "masterXmlText": xml_text
            if master
            else shared_master[1]
            if shared_master
            else None,
            "translationStatus": "master"
            if master
            else translation_status or "not_attempted",
            "translatedFormula": translated_formula if not master else None,
            "translationMethod": "openpyxl.Translator"
            if (not master and translated_formula is not None)
            else None,
        }
        if not master:
            provenance["status"] = (
                "shared_translation"
                if (translation_status == "translated")
                else "unavailable"
            )
    result.update(
        sourceFormula=source_formula,
        canonicalFormula=canonical_formula,
        formulaProvenance=provenance,
    )
    return result


def _cells(archive: zipfile.ZipFile, part: str):
    """Stream complete cells while detaching cell and row objects."""
    with archive.open(part) as source:
        stack = []
        row_number = 0
        previous_column = 0
        for event, element in ET.iterparse(source, events=("start", "end")):
            if event == "start":
                stack.append(element)
                if element.tag.rsplit("}", 1)[-1] == "row":
                    try:
                        row_number = int(element.get("r", str(row_number + 1)))
                    except ValueError:
                        row_number = 0
                    previous_column = 0
                continue
            tag = element.tag.rsplit("}", 1)[-1]
            if tag == "c":
                address = element.get("r")
                if address is None and row_number > 0:
                    address = get_column_letter(previous_column + 1) + str(row_number)
                if address:
                    try:
                        previous_column = coordinate_to_tuple(address)[1]
                    except ValueError:
                        address = None
                yield element, address
                if len(stack) > 1:
                    stack[-2].remove(element)
                element.clear()
            elif tag == "row":
                if len(stack) > 1:
                    stack[-2].remove(element)
                element.clear()
            stack.pop()


def _worksheet_parts(archive: zipfile.ZipFile):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relations = {
        element.get("Id"): element
        for element in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    }
    for sheet in workbook.findall("{*}sheets/{*}sheet"):
        rid = next(
            (v for k, v in sheet.attrib.items() if k.rsplit("}", 1)[-1] == "id"), None
        )
        relation = relations.get(rid)
        if relation is None or relation.get("TargetMode") == "External":
            continue
        if not relation.get("Type", "").endswith("/worksheet"):
            continue
        target = relation.get("Target", "")
        part = posixpath.normpath(
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.join("xl", target)
        )
        yield sheet.get("name"), part


def enrich_graph_sources(
    data: bytes, nodes: Iterable[dict[str, Any]], warnings: list[str]
) -> None:
    """Attach source evidence to retained cells/samples, never rewrite formula."""
    digest = hashlib.sha256(data).hexdigest()
    wanted: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for node in nodes:
        if node.get("kind") not in {"cell", "group", "input"}:
            continue
        sheet = node.get("sheet")
        if not isinstance(sheet, str):
            continue
        records = [node, *(node.get("samples") or []), *(node.get("values") or [])]
        for record in records:
            cell = record.get("addr", record.get("cell"))
            if not isinstance(cell, str) or not re.fullmatch(
                r"[A-Z]{1,3}[1-9]\d*", cell
            ):
                continue
            wanted.setdefault((sheet, cell), []).append(record)
            if record.get("formula"):
                record["canonicalFormula"] = record["formula"]
            if record.get("steps") is not None:
                record["decompositionFormulaBasis"] = "canonicalFormula"
    if not wanted:
        return
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            for sheet, part in _worksheet_parts(archive):
                addresses = {cell for name, cell in wanted if name == sheet}
                if not addresses:
                    continue
                captured = {}
                shared_ids = set()
                for element, address in _cells(archive, part):
                    if address not in addresses:
                        continue
                    captured[address] = ET.fromstring(ET.tostring(element))
                    formula = element.find("{*}f")
                    if (
                        formula is not None
                        and formula.get("t") == "shared"
                        and not formula.text
                    ):
                        shared_ids.add(formula.get("si"))
                masters = {}
                if shared_ids:
                    for element, address in _cells(archive, part):
                        formula = element.find("{*}f")
                        if (
                            formula is not None
                            and formula.get("t") == "shared"
                            and formula.get("si") in shared_ids
                            and formula.text
                        ):
                            masters[formula.get("si")] = (
                                address,
                                formula.text,
                            )
                for address in addresses:
                    element = captured.get(address, ET.Element("c"))
                    formula = element.find("{*}f")
                    master = (
                        masters.get(formula.get("si")) if formula is not None else None
                    )
                    translation = None
                    translated = None
                    if (
                        formula is not None
                        and formula.get("t") == "shared"
                        and not formula.text
                    ):
                        translation = "missing_master"
                        if master:
                            try:
                                translated = Translator(
                                    "=" + master[1], origin=master[0]
                                ).translate_formula(address)
                                translation = "translated"
                            except Exception:
                                translation = "failed"
                    for record in wanted[(sheet, address)]:
                        canonical = record.get(
                            "formula", record.get("canonicalFormula")
                        )
                        record.update(
                            cell_source_fields(
                                element,
                                workbook_sha256=digest,
                                worksheet_part=part,
                                sheet=sheet,
                                cell=address,
                                canonical_formula=canonical,
                                canonical_origin="engine"
                                if canonical
                                else "unavailable",
                                shared_master=master,
                                translation_status=translation,
                                translated_formula=translated,
                            )
                        )
    except (zipfile.BadZipFile, KeyError, OSError, ET.ParseError) as exc:
        warnings.append(
            f"Stored formula evidence unavailable: {type(exc).__name__}: {exc}"
        )
        for records in wanted.values():
            for record in records:
                if record.get("formula"):
                    record.setdefault("sourceFormula", None)
                    record.setdefault(
                        "formulaProvenance",
                        {
                            "status": "unavailable",
                            "workbookSha256": digest,
                            "reason": "OOXML evidence could not be read",
                            "canonicalOrigin": "engine",
                        },
                    )
