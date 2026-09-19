"""Bounded source facts for documentation, independent of formula evaluation."""

from __future__ import annotations

import io
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

import formualizer as fz

MAX_NAME_XML_BYTES = 2 * 1024 * 1024
MAX_NAME_RECORDS = 200
MAX_NAME_EXPRESSION_CHARS = 1000
MAX_FORMULA_FACT_CHARS = 8192


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


def formula_evidence(formula: str | None) -> dict[str, Any]:
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
    local_references_omitted = False
    pending = [(ast, False)]
    while pending:
        item, local_scope = pending.pop()
        if isinstance(item, list):
            pending.extend((child, local_scope) for child in reversed(item))
        elif isinstance(item, dict):
            if item.get("node_type") == "Reference" and not local_scope:
                references.add(item["reference"])
            if item.get("node_type") == "Function":
                name = item.get("name", "").upper().removeprefix("_XLFN.")
                if name in {"LET", "LAMBDA"}:
                    # Do not confuse bindings with workbook names. The graph
                    # retains the actual precedents; this optional name lookup
                    # conservatively omits references under lexical bindings.
                    local_scope = True
                    local_references_omitted = True
                args = item.get("args", [])
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
    return {
        "status": "parsed",
        "references": sorted(references),
        "local_scope_references_omitted": local_references_omitted,
        "function_selectors": selectors[:20],
        "function_selectors_omitted": max(0, len(selectors) - 20),
    }


def _literal_integer(arg: dict) -> int | None:
    value = arg.get("value") if arg.get("node_type") == "Literal" else None
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        if -1000 <= value <= 1000 and value == int(value):
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
