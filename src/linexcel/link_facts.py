"""External reference grammar and immutable OOXML facts, without resolution.

Only external A1 cells/rectangles are supported. Tokenized external names,
tables, row/column selections and 3-D references remain explicit facts.
Paths are evidence, never filesystem instructions.
"""

from __future__ import annotations

import re

from linexcel.external import parse_external_refs, read_external_links
from linexcel.refs import parse_ref
from linexcel.structured_refs import formula_tokens

_BOOK = re.compile(r"\[([^\]]+)\]")
_A1 = re.compile(r"\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?")


def reference_facts(formula, books):
    """Retain every external operand, including forms we cannot traverse."""
    from linexcel.lazy import _json

    try:
        tokens = formula_tokens(formula)
        operands = [
            t.value for t in tokens if t.type == "OPERAND" and t.subtype == "RANGE"
        ]
    except Exception:
        # A malformed formula still discloses external workbook markers.
        tokens = []
        operands = [formula] if _BOOK.search(formula) else []
    dynamic = any(
        t.type == "FUNC"
        and t.subtype == "OPEN"
        and t.value[:-1].upper().removeprefix("_XLFN.") in {"INDIRECT", "OFFSET"}
        for t in tokens
    )
    if dynamic:
        # A quoted external operand in INDIRECT is evidence of an unsupported
        # dynamic reference, never an executable static reference.
        operands.extend(
            t.value[1:-1].replace('""', '"')
            for t in tokens
            if t.type == "OPERAND" and t.subtype == "TEXT"
        )
    facts = []
    for operand in operands:
        marker = _BOOK.search(operand)
        if not marker:
            continue
        prefix = operand[: marker.start()]
        if "!" not in operand or not (
            prefix in {"", "'"} or prefix.endswith(("/", "\\"))
        ):
            continue
        key = marker[1]
        # A local Table[Column] is not an external reference.
        if (
            key not in books
            and not key.isdigit()
            and not re.search(r"\.(?:xlsx|xlsm|xlsb|xlam|xla|xltx|xltm)$", key, re.I)
        ):
            continue
        book = books.get(key)
        parsed = parse_external_refs(operand)
        ref = parsed[0] if len(parsed) == 1 and parsed[0].text == operand else None
        sheet = ref.sheet if ref else None
        address = ref.cell.upper() if ref else None
        rect = parse_ref(address, default_sheet=sheet) if address else None
        supported = bool(
            ref
            and _A1.fullmatch(ref.cell)
            and rect
            and sheet
            and ":" not in sheet
            and not dynamic
        )
        subtype = (
            "a1"
            if supported
            else (
                "dynamic"
                if dynamic
                else "three_dimensional"
                if sheet and ":" in sheet
                else "structured"
                if "[" in operand[marker.end() :]
                else "whole_row_or_column"
                if ":" in operand.rsplit("!", 1)[-1]
                else "defined_name_or_invalid"
            )
        )
        target = (
            book.relationship_path or book.target
            if book
            else ((ref.directory if ref else "") + key)
        )
        cached = []
        if book and rect:
            for (s, r, c), value in book.cached.items():
                if (
                    s.lower() == (sheet or "").lower()
                    and rect.r1 <= r <= rect.r2
                    and rect.c1 <= c <= rect.c2
                ):
                    cached.append({"row": r, "column": c, "value": _json(value)})
        facts.append(
            {
                "bookKey": key,
                "relationshipPath": target,
                "filename": book.name if book else key,
                "reference": operand,
                "sheet": sheet,
                "address": address,
                "bounds": [rect.r1, rect.c1, rect.r2, rect.c2] if supported else None,
                "supported": supported,
                "referenceType": subtype,
                "cached": cached,
            }
        )
    return facts


def index_facts(path, store, metadata):
    from linexcel.lazy_store import _dump, _unpack

    db = store.db
    db.executescript("""
        CREATE TABLE external_facts(node TEXT NOT NULL,ordinal INTEGER NOT NULL,
            data TEXT NOT NULL,PRIMARY KEY(node,ordinal)) WITHOUT ROWID;
        CREATE TABLE lineage_ranges(node TEXT PRIMARY KEY,sheet TEXT NOT NULL,
            r1 INTEGER,c1 INTEGER,r2 INTEGER,c2 INTEGER);
        CREATE INDEX lineage_range_sheet ON lineage_ranges(sheet,r1,r2);
    """)
    books = read_external_links(path)
    metadata["externalFactsVersion"] = 1
    from linexcel.powerquery import read_destinations

    metadata["powerQueryDestinations"] = [
        {"query": name, "sheet": dest.sheet, "ref": dest.ref, "table": dest.table}
        for name, destinations in read_destinations(path).items()
        for dest in destinations
    ]
    metadata["externalBooks"] = [
        {
            "bookKey": b.key,
            "relationshipPath": b.relationship_path,
            "filename": b.name,
            "sheets": b.sheets,
        }
        for b in books.values()
    ]
    for nid, data in db.execute(
        "SELECT id,data FROM nodes WHERE formula IS NOT NULL AND instr(formula,'[')>0"
    ):
        node = _unpack(data)
        if node.get("external"):
            continue
        for ordinal, fact in enumerate(reference_facts(node["formula"], books)):
            # Match existing external placeholders, so traversal never invents
            # a local edge for their disconnected materialized cache nodes.
            fact["placeholder"] = node.get("reference_bindings", {}).get(
                fact["reference"]
            )
            db.execute(
                "INSERT INTO external_facts VALUES (?,?,?)", (nid, ordinal, _dump(fact))
            )
    for nid, data in db.execute("SELECT id,data FROM nodes WHERE kind='range'"):
        node = _unpack(data)
        if not node.get("external") and node.get("bounds") and node.get("sheet"):
            db.execute(
                "INSERT INTO lineage_ranges VALUES (?,?,?,?,?,?)",
                (nid, node["sheet"].lower(), *node["bounds"]),
            )
    metadata["externalReferenceCount"] = db.execute(
        "SELECT count(*) FROM external_facts"
    ).fetchone()[0]


def facts_available(index):
    return index.meta.get("externalFactsVersion") == 1
