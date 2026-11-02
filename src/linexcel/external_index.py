"""Sparse materialization of dependencies from explicitly supplied workbooks."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from pathlib import Path, PureWindowsPath
from typing import Any

from linexcel.external import _EXTERNAL_REF_RE
from linexcel.refs import a1, parse_ref, split_sheet_prefix
from linexcel.structured_refs import resolve_table_reference


class ExternalSources:
    def __init__(self, store, refs_dir, temporary, links, date_system):
        self.store = store
        self.refs_dir = refs_dir
        self.temporary = Path(temporary)
        self.links = links
        self.date_system = date_system
        self.books = {}

    def close(self):
        for index, *_ in self.books.values():
            index.close()

    def load(self, filename):
        from linexcel.lazy_store import (
            WorkbookIndex,
            _read_source,
            _schema,
            file_sha256,
        )

        if filename in self.books:
            return self.books[filename]
        if self.refs_dir is None or PureWindowsPath(filename).name != filename:
            raise ValueError("External workbook was not supplied: " + filename)
        source = self.refs_dir / filename
        if (
            not source.is_file()
            or source.is_symlink()
            or source.resolve().parent != self.refs_dir.resolve()
        ):
            raise ValueError("External workbook was not supplied: " + filename)
        index = WorkbookIndex(
            self.temporary / f"{len(self.books)}.sqlite", writable=True
        )
        self.books[filename] = (index, {}, {}, {}, "")
        _schema(index.db)
        metadata, names, links = _read_source(source, index)
        record = index, metadata, names, links, file_sha256(source)
        self.books[filename] = record
        return record

    @staticmethod
    def namespace(filename, sheet):
        return f"[{filename}]{sheet}"

    def copy_member(self, filename, sheet, coordinate, reference):
        from linexcel.lazy import _id

        nid = f"external:{filename}:{_id(sheet, coordinate)}"
        if nid in self.store:
            return nid
        index, metadata, _names, _links, digest = self.load(filename)
        rect = parse_ref(coordinate, default_sheet=sheet)
        assert rect is not None
        member = index.get(_id(sheet, coordinate))
        node: dict[str, Any] = (
            deepcopy(member)
            if member is not None
            else {
                "kind": "input",
                "row": rect.r1,
                "column": rect.c1,
                "cell": coordinate,
                "formula": None,
                "value": None,
                "cached_value": None,
                "cache_status": "blank",
                "sourceDataType": "n",
                "cachedValueKind": "scalar",
                "diagnostics": [],
            }
        )
        node.update(
            id=nid,
            external=True,
            sheet=self.namespace(filename, sheet),
            sourceSheet=sheet,
            sourceFile=filename,
            sourceSha256=digest,
            sourceCalculationSettings=metadata.get("calculationSettings"),
            reference=reference,
            dependencies=[],
            provenance="external_formula_cache"
            if node.get("formula")
            else "reference_input",
        )
        if any(
            s["name"] == sheet and s["status"] != "complete"
            for s in metadata["sheetDetails"]
        ):
            node["diagnostics"].append("External source sheet import is incomplete.")
        if index.db.execute(
            "SELECT 1 FROM arrays WHERE sheet=? AND ? BETWEEN r1 AND r2 "
            "AND ? BETWEEN c1 AND c2 LIMIT 1",
            (sheet, rect.r1, rect.c1),
        ).fetchone():
            node["diagnostics"].append(
                "External array result semantics are unsupported."
            )
        if metadata["dateSystem"] != self.date_system:
            from openpyxl.styles.numbers import is_date_format

            if node.get("formula") or is_date_format(
                node.get("numberFormat", "General")
            ):
                node["diagnostics"].append(
                    "External formula/date uses a different workbook calendar; "
                    "cross-calendar calculation is not verified."
                )
        self.store.put(node, only_if_absent=True)
        return nid

    def rectangle(self, filename, rect, reference):
        index, metadata, _names, _links, _digest = self.load(filename)
        sheets = {s.casefold(): s for s in metadata.get("sheets", [])}
        if rect.sheet is None or rect.sheet.casefold() not in sheets:
            raise ValueError("External source sheet is unavailable.")
        rect = replace(rect, sheet=sheets[rect.sheet.casefold()])
        if rect.ncells == 1:
            return self.copy_member(
                filename, rect.sheet, a1(rect.r1, rect.c1), reference
            )
        nid = f"external:{filename}:{rect.to_a1()}"
        if nid in self.store:
            return nid
        node = {
            "id": nid,
            "kind": "range",
            "external": True,
            "sheet": self.namespace(filename, rect.sheet),
            "sourceSheet": rect.sheet,
            "sourceFile": filename,
            "reference": reference,
            "bounds": [rect.r1, rect.c1, rect.r2, rect.c2],
            "cell_count": rect.ncells,
            "dependencies": [],
            "diagnostics": [],
        }
        if any(
            s["name"] == rect.sheet and s["status"] != "complete"
            for s in metadata["sheetDetails"]
        ):
            node["diagnostics"].append("External source sheet import is incomplete.")
        self.store.put(node, only_if_absent=True)
        for member_id in index.range_members(rect.sheet, node["bounds"]):
            self.copy_member(filename, rect.sheet, index[member_id]["cell"], reference)
        return nid

    def explicit(self, text, source_node):
        match = _EXTERNAL_REF_RE.fullmatch(text)
        if not match:
            return None
        links = self.links
        if source_node.get("external"):
            _, _, _, links, _ = self.load(source_node["sourceFile"])
        filename = links.get(match.group("book"), match.group("book"))
        sheet = match.group("sheet").replace("''", "'")
        rect = parse_ref(match.group("cell"), default_sheet=sheet)
        if rect is None:
            return None
        try:
            return self.rectangle(filename, rect, text)
        except (ValueError, OSError, KeyError) as exc:
            nid = f"external-unresolved:{filename}:{rect.to_a1()}"
            self.store.put(
                {
                    "id": nid,
                    "kind": "unresolved",
                    "external": True,
                    "sourceFile": filename,
                    "sourceSheet": sheet,
                    "sheet": self.namespace(filename, sheet),
                    "reference": text,
                    "dependencies": [],
                    "diagnostics": ["Unresolved external reference: " + str(exc)],
                },
                only_if_absent=True,
            )
            return nid

    def local(self, text, source_node):
        """Bind an external formula's ordinary references in its source workbook."""
        filename = source_node["sourceFile"]
        index, metadata, names, _links, digest = self.load(filename)
        prefix, body = split_sheet_prefix(text)
        sheet = source_node.get("sourceSheet")
        lookup = {s.casefold(): s for s in metadata["sheets"]}
        scope = lookup.get(prefix.casefold()) if prefix is not None else sheet
        name_id = (
            names.get((scope, body.casefold())) or names.get((None, body.casefold()))
            if prefix is None or scope is not None
            else None
        )
        if name_id:
            nid = f"external:{filename}:{name_id}"
            if nid not in self.store:
                named = deepcopy(index[name_id])
                named.update(
                    id=nid,
                    external=True,
                    sourceFile=filename,
                    sourceSheet=named.get("sheet"),
                    sourceSha256=digest,
                    sheet=self.namespace(filename, named.get("sheet"))
                    if named.get("sheet")
                    else None,
                )
                self.store.put(named, only_if_absent=True)
            return nid
        selection = resolve_table_reference(
            index, text, sheet, source_node.get("row"), source_node.get("column")
        )
        rect = (
            selection.rectangle
            if selection is not None
            else parse_ref(text, default_sheet=sheet)
        )
        if rect is None:
            raise ValueError("Unresolved external source dependency: " + text)
        if selection is not None:
            source_node.setdefault("externalTableReferences", []).append(text)
        return self.rectangle(filename, rect, text)
