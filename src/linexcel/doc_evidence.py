"""Bounded source facts for documentation, independent of formula evaluation."""

from __future__ import annotations

import hashlib
import io
import re
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

import formualizer as fz

from linexcel.refs import split_sheet_prefix

MAX_NAME_XML_BYTES = 2 * 1024 * 1024
MAX_NAME_RECORDS = 200
MAX_NAME_EXPRESSION_CHARS = 1000
MAX_FORMULA_FACT_CHARS = 8192


def formula_text_evidence(node: dict[str, Any]) -> dict[str, Any]:
    """Measure stored characters separately from display and parser spelling.

    No expression equivalence is inferred. Older records without XML evidence
    remain unknown even when a field happens to begin with an equals sign.
    """
    provenance = node.get("formulaProvenance") or {}
    xml_text = provenance.get("xmlText")
    source = node.get("sourceFormula")
    canonical = node.get("canonicalFormula", node.get("formula"))
    xml_available = isinstance(xml_text, str)
    source_available = isinstance(source, str)
    canonical_available = isinstance(canonical, str)
    comparable = xml_available and bool(xml_text) and source_available
    return {
        "stored_xml_character_count": len(xml_text) if xml_available else None,
        "presented_source_character_count": len(source) if source_available else None,
        "canonical_character_count": len(canonical) if canonical_available else None,
        "source_to_xml_relation": (
            "presentation_equals_prefix_plus_decoded_xml"
            if comparable and source == "=" + xml_text
            else "inconsistent"
            if comparable
            else "not_comparable"
        ),
        "added_prefix": "=" if comparable and source == "=" + xml_text else None,
        "canonical_to_presented_source": (
            "literal_equal"
            if canonical_available and source_available and canonical == source
            else "literal_different"
            if canonical_available and source_available
            else "not_comparable"
        ),
        "scope": "decoded_source_record_not_xml_bytes_or_semantic_proof",
    }


def value_comparison_evidence(node: dict[str, Any]) -> dict[str, Any]:
    """Compare supplied JSON values without parsing or normalizing string text."""
    value, cached = node.get("value"), node.get("cachedValue")
    available = "value" in node and "cachedValue" in node
    both_strings = available and isinstance(value, str) and isinstance(cached, str)
    any_string = available and (isinstance(value, str) or isinstance(cached, str))
    return {
        "scope": "displayed_value_vs_saved_cache",
        "displayed_json_type": _json_type(value) if "value" in node else "unknown",
        "cached_json_type": _json_type(cached) if "cachedValue" in node else "unknown",
        "literal_text_equal": value == cached if both_strings else None,
        "numeric_equivalence": "not_established_from_strings"
        if any_string
        else "not_assessed",
        "comparison_label": node.get("cachedAgreement"),
        "comparison_label_scope": "formatting_heuristic_not_numeric_equivalence",
    }


def _json_type(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "string"
    if isinstance(value, list):
        return "array"
    return "object" if isinstance(value, dict) else "unknown"


def vba_excerpt_evidence(node: dict[str, Any], *, limit: int = 2500) -> dict[str, Any]:
    """Keep full-source, graph-preview and sent-excerpt bounds independent.

    Hashes of decoded text use UTF-8; character counts use Python characters,
    never bytes or line counts. Full-source size is unknown without its explicit
    provenance. A supplied full-source hash is not the graph/excerpt hash.
    """
    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 0:
        raise ValueError("VBA excerpt limit must be a nonnegative integer")
    code = node.get("code")
    graph_available = isinstance(code, str)
    excerpt = code[:limit] if graph_available else None
    graph_count = len(code) if graph_available else None
    excerpt_count = len(excerpt) if excerpt is not None else None
    provenance = node.get("vbaProvenance") or {}
    original = provenance.get("totalCodeCharacters")
    full_count = (
        original
        if isinstance(original, int)
        and not isinstance(original, bool)
        and original >= 0
        else None
    )
    consistent = (
        full_count >= graph_count
        and (
            provenance.get("codeTruncated") is None
            or provenance["codeTruncated"] == (full_count > graph_count)
        )
        if full_count is not None and graph_count is not None
        else None
    )
    return {
        "code": excerpt,
        "code_status": "supplied_excerpt" if graph_available else "unavailable",
        "code_omitted_characters": graph_count - excerpt_count
        if graph_count is not None and excerpt_count is not None
        else None,
        "code_bounds": {
            "unit": "decoded_source_characters",
            "full_source_characters": full_count,
            "graph_characters": graph_count,
            "dossier_characters": excerpt_count,
            "omitted_full_source_to_graph": full_count - graph_count
            if consistent is True and full_count is not None and graph_count is not None
            else None,
            "omitted_graph_to_dossier": graph_count - excerpt_count
            if graph_count is not None and excerpt_count is not None
            else None,
            "omitted_full_source_to_dossier": full_count - excerpt_count
            if consistent is True
            and full_count is not None
            and excerpt_count is not None
            else None,
            "source_and_graph_size_consistent": consistent,
        },
        "code_hashes": {
            "full_source_sha256": provenance.get("procedureSha256"),
            "graph_sha256": hashlib.sha256(code.encode("utf-8")).hexdigest()
            if graph_available
            else None,
            "dossier_sha256": hashlib.sha256(excerpt.encode("utf-8")).hexdigest()
            if excerpt is not None
            else None,
        },
        "code_bound_semantics": (
            "Counts measure distinct full/graph/dossier decoded strings. "
            "Only the supplied excerpt is readable; source hashes, spans and "
            "omission counts do not expose omitted text or prove execution."
        ),
    }


def read_defined_name_evidence(data: bytes) -> dict[str, Any]:
    """Keep constants and local scope that reference-only name nodes cannot show.

    No expression is evaluated. Oversized or unavailable metadata is explicitly
    incomplete, so missing evidence cannot be interpreted as a missing name.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            part = package.getinfo("xl/workbook.xml")
            if part.file_size > MAX_NAME_XML_BYTES:
                return {"status": "not_inspected", "reason": "XML size limit"}
            payload = package.read(part)
        # This metadata reader has no reason to accept entity definitions.
        declarations = payload.replace(b"\x00", b"")
        if b"<!DOCTYPE" in declarations or b"<!ENTITY" in declarations:
            return {"status": "not_inspected", "reason": "XML declarations"}
        root = ET.fromstring(payload)
        ns = root.tag.rsplit("}", 1)[0] + "}" if "}" in root.tag else ""
        sheets = [s.get("name") for s in root.findall(f"{ns}sheets/{ns}sheet")]
        definitions = root.findall(f"{ns}definedNames/{ns}definedName")
        records = []
        for item in definitions[:MAX_NAME_RECORDS]:
            scope_id = item.get("localSheetId")
            scope_sheet = None
            scope = "workbook" if scope_id is None else "unknown"
            if scope_id is not None:
                try:
                    index = int(scope_id)
                    if 0 <= index < len(sheets):
                        scope_sheet = sheets[index]
                        scope = "sheet"
                except ValueError:
                    pass
            expression = item.text or ""
            records.append(
                {
                    "name": item.get("name"),
                    "scope": scope,
                    "scope_sheet": scope_sheet,
                    "expression": expression
                    if len(expression) <= MAX_NAME_EXPRESSION_CHARS
                    else {"omitted": "length limit", "characters": len(expression)},
                    "expression_evaluated": False,
                }
            )
        return {
            "status": "complete" if len(records) == len(definitions) else "partial",
            "total": len(definitions),
            "omitted": len(definitions) - len(records),
            "definitions": records,
            "sheets": sheets,
        }
    except (KeyError, ValueError, OSError, zipfile.BadZipFile, ET.ParseError):
        return {"status": "not_inspected", "reason": "unavailable workbook metadata"}


# Literal selector definitions, not claims about the engine's result.
# https://support.microsoft.com/en-us/excel/functions/aggregate-function
_AGGREGATES = (
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "MAX",
    "MIN",
    "PRODUCT",
    "STDEV.S",
    "STDEV.P",
    "SUM",
    "VAR.S",
    "VAR.P",
    "MEDIAN",
    "MODE.SNGL",
    "LARGE",
    "SMALL",
    "PERCENTILE.INC",
    "QUARTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE.EXC",
)


def reference_anchor_evidence(reference: str) -> dict[str, Any] | None:
    """Literal A1 anchor flags, never inferred from a reference's computed value."""
    _, body = split_sheet_prefix(reference)
    endpoints = body.split(":")
    if len(endpoints) not in (1, 2):
        return None
    anchors = []
    for endpoint in endpoints:
        match = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)", endpoint)
        if match is None:
            return None
        column_absolute, row_absolute = bool(match[1]), bool(match[3])
        anchors.append(
            {
                "address": endpoint,
                "column_absolute": column_absolute,
                "row_absolute": row_absolute,
                "mode": "absolute"
                if column_absolute and row_absolute
                else "mixed"
                if column_absolute or row_absolute
                else "relative",
            }
        )
    return {"reference": reference, "endpoints": anchors}


def formula_evidence(
    formula: str | None, *, sheet: str | None = None, address: str | None = None
) -> dict[str, Any]:
    """Describe selected AST facts without guessing computed selector values."""
    if not formula:
        return {"status": "no_formula", "references": [], "function_selectors": []}
    if len(formula) > MAX_FORMULA_FACT_CHARS:
        return {"status": "not_inspected", "reason": "formula length limit"}
    # parse().to_dict() itself can overflow the native stack on a long flat
    # operator chain. Do not call a parser-based depth guard here. This cheap
    # conservative bound may also omit punctuation in strings; only optional
    # documentation facts are omitted, never the formula or its calculation.
    if sum(char in "=+-*/^&(),;<>:%{}" for char in formula) > 128:
        return {"status": "not_inspected", "reason": "formula complexity limit"}
    try:
        ast = fz.parse(formula if formula.startswith("=") else "=" + formula).to_dict()
    except Exception:
        return {"status": "not_inspected", "reason": "formula parse failure"}
    references: set[str] = set()
    selectors = []
    functions: set[str] = set()
    lookups = []
    local_references_omitted = False
    reference_operators = []
    pending = [(ast, False)]
    while pending:
        item, local_scope = pending.pop()
        if isinstance(item, list):
            pending.extend((child, local_scope) for child in reversed(item))
        elif isinstance(item, dict):
            if item.get("node_type") == "BinaryOp" and item.get("operator") in {
                " ",
                ",",
            }:
                reference_operators.append(
                    {
                        "operator": item["operator"],
                        "meaning": "intersection"
                        if item["operator"] == " "
                        else "union",
                        "left_reference": item.get("left", {}).get("reference"),
                        "right_reference": item.get("right", {}).get("reference"),
                        "scope": (
                            "parsed reference operator; "
                            "no evaluated range/result implied"
                        ),
                    }
                )
            if item.get("node_type") == "Reference" and not local_scope:
                references.add(item["reference"])
            if item.get("node_type") == "Function":
                name = item.get("name", "").upper().removeprefix("_XLFN.")
                functions.add(name)
                if name in {"LET", "LAMBDA"}:
                    # Do not confuse bindings with workbook names. The graph
                    # retains the actual precedents; this optional name lookup
                    # conservatively omits references under lexical bindings.
                    local_scope = True
                    local_references_omitted = True
                args = item.get("args", [])
                if name in {"VLOOKUP", "HLOOKUP"} and len(args) >= 3:
                    index = _literal_integer(args[2], max_abs=1_048_576)
                    table = (
                        args[1].get("reference")
                        if args[1].get("node_type") == "Reference"
                        else None
                    )
                    lookups.append(
                        {
                            "function": name,
                            "table_reference": table,
                            "one_based_index": index,
                            "offset_from_first": index - 1
                            if index is not None and index >= 1
                            else None,
                            "axis": "columns" if name == "VLOOKUP" else "rows",
                            "basis": (
                                "literal AST index; offset is index minus one, "
                                "not an evaluated lookup result"
                            ),
                        }
                    )
                if name == "AGGREGATE" and len(args) >= 2:
                    selector, option = (
                        _literal_integer(args[0]),
                        _literal_integer(args[1]),
                    )
                    fact: dict[str, Any] = {
                        "function": name,
                        "function_num": selector,
                        "options": option,
                        "basis": "literal AST arguments only; null means undetermined",
                    }
                    if selector is not None and 1 <= selector <= len(_AGGREGATES):
                        fact["operation"] = _AGGREGATES[selector - 1]
                    if option is not None and 0 <= option <= 7:
                        fact["ignore_errors"] = option in (2, 3, 6, 7)
                        fact["ignore_hidden_rows_requested"] = option in (1, 3, 5, 7)
                        fact["ignore_nested_aggregates_requested"] = option in (
                            0,
                            1,
                            2,
                            3,
                        )
                        fact["caveat"] = (
                            "Calculated array arguments do not ignore hidden rows or "
                            "nested aggregates. These are Excel selector meanings, "
                            "not verification of engine behavior."
                        )
                    selectors.append(fact)
            pending.extend(
                (child, local_scope) for child in reversed(list(item.values()))
            )
    spans = [
        fact
        for reference in sorted(references)
        if (fact := three_dimensional_reference_evidence(reference)) is not None
    ]
    anchors = [
        fact
        for reference in sorted(references)
        if (fact := reference_anchor_evidence(reference)) is not None
    ]
    return {
        "status": "parsed",
        "references": sorted(references),
        "reference_anchors": anchors[:40],
        "reference_anchors_omitted": max(0, len(anchors) - 40),
        "reference_anchor_scope": (
            "literal A1 dollar markers; named and structured references "
            "are not classified as absolute or relative cell references"
        ),
        "local_scope_references_omitted": local_references_omitted,
        "function_selectors": selectors[:20],
        "function_selectors_omitted": max(0, len(selectors) - 20),
        "functions_observed": sorted(functions),
        "function_inventory_scope": (
            "parsed syntax, not engine support or error-cause evidence"
        ),
        "lookup_offsets": lookups[:20],
        "lookup_offsets_omitted": max(0, len(lookups) - 20),
        "reference_operators": reference_operators[:20],
        "reference_operators_omitted": max(0, len(reference_operators) - 20),
        **(
            {
                "three_dimensional_references": spans[:20],
                "three_dimensional_references_omitted": max(0, len(spans) - 20),
            }
            if spans
            else {}
        ),
        "literal_self_reference": _literal_self_reference(references, sheet, address)
        if not local_references_omitted
        else None,
        "self_reference_scope": (
            "literal single-cell references only; "
            "not a cycle or dynamic-reference proof"
        ),
    }


def three_dimensional_reference_evidence(reference: Any) -> dict[str, Any] | None:
    """Identify a literal sheet-span qualifier, without resolving its sheets.

    Used on AST reference tokens, never raw formula text/string literals. The
    optional decomposition caller can also use it on a recorded range token.
    A colon inside an external workbook/path is not a sheet-span separator.
    Neither sheet order nor a total cell count follows from this token alone.
    """
    if not isinstance(reference, str):
        return None
    qualifier, address = split_sheet_prefix(reference)
    if qualifier is None:
        return None
    sheet_qualifier = qualifier
    external = "[" in qualifier or "]" in qualifier
    if external:
        if qualifier.count("[") != 1 or qualifier.count("]") != 1:
            return None
        opening, closing = qualifier.index("["), qualifier.index("]")
        if opening >= closing:
            return None
        sheet_qualifier = qualifier[closing + 1 :]
    endpoints = sheet_qualifier.split(":")
    if len(endpoints) != 2 or not all(endpoints) or not address:
        return None
    return {
        "reference": reference,
        "spans_sheets": True,
        "first_sheet_token": endpoints[0],
        "last_sheet_token": endpoints[1],
        "address_per_sheet": address,
        "external_workbook_qualifier": external,
        "sheet_count": None,
        "total_cells": None,
        "scope": (
            "literal 3D sheet-span syntax only; an A1 address is per sheet, "
            "not proof of one scalar input. Sheet order/count, total cells, "
            "resolution and evaluated values are not established."
        ),
    }


def _literal_self_reference(
    references: set[str], sheet: str | None, address: str | None
) -> bool | None:
    if not sheet or not address:
        return None
    for reference in references:
        qualifier, separator, cell = reference.rpartition("!")
        if (
            separator
            and qualifier.strip("'").replace("''", "'").casefold() != sheet.casefold()
        ):
            continue
        if cell.replace("$", "").casefold() == address.replace("$", "").casefold():
            return True
    return False


def _literal_integer(arg: dict, *, max_abs: int = 1000) -> int | None:
    value = arg.get("value") if arg.get("node_type") == "Literal" else None
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        if -max_abs <= value <= max_abs and value == int(value):
            return int(value)
    return None


def relevant_names(evidence: dict, references: list[str], sheet: str | None) -> dict:
    """Return matching source definitions, preserving local/global ambiguity."""
    refs = []
    for reference in references:
        qualifier, separator, name = reference.rpartition("!")
        if not separator:
            refs.append((None, reference.casefold()))
        else:
            if qualifier.startswith("'") and qualifier.endswith("'"):
                qualifier = qualifier[1:-1].replace("''", "'")
            refs.append((qualifier.casefold(), name.casefold()))
    sheets = {name.casefold() for name in evidence.get("sheets", []) if name}
    definitions = []
    for item in evidence.get("definitions", []):
        name = (item.get("name") or "").casefold()
        scope_sheet = (item.get("scope_sheet") or "").casefold()
        for qualifier, ref in refs:
            if ref != name:
                continue
            active_sheet = (
                qualifier if qualifier is not None else (sheet or "").casefold()
            )
            applies = (
                item.get("scope") == "workbook"
                and (qualifier is None or qualifier in sheets)
            ) or (item.get("scope") == "sheet" and scope_sheet == active_sheet)
            if applies:
                definitions.append(item)
                break
    return {
        "metadata_status": evidence.get("status", "not_inspected"),
        "selection_scope": "names_referenced_by_this_formula_only",
        "workbook_definition_count": evidence.get("total"),
        "definitions": definitions,
        "interpretation": (
            "This filtered list is not the workbook name inventory or the names "
            "used by neighboring formulas. An empty list does not mean the "
            "workbook has no defined names. "
            "A sheet-local definition takes precedence over the same workbook name. "
            "A recorded definition does not prove engine support; a missing graph "
            "edge or #NAME? result does not establish that the source name is absent."
        ),
    }
