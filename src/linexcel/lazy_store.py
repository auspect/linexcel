"""Disk-backed, sparse OOXML index for the interactive lazy viewer.

ZIP inflation and XML parsing use the standard library's native zlib/Expat
implementations; SQLite owns the indexes and the bounded page cache. Only one
stored cell is decoded at a time. In particular shared strings, copied formula
members and worksheet row objects never accumulate in Python memory.
"""

from __future__ import annotations

import hashlib
import heapq
import json
import math
import os
import posixpath
import re
import sqlite3
import tempfile
import zipfile
import zlib
from collections.abc import Mapping
from dataclasses import replace
from functools import partial
from pathlib import Path, PureWindowsPath
from typing import Any
from urllib.parse import quote, unquote
from xml.etree import ElementTree as ET

from openpyxl.formula.translate import Translator
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
    from_excel,
    from_ISO8601,
    to_excel,
)

from linexcel.cell_grid import (
    GRID_COLUMNS,
    GRID_SCHEMA,
    GridReader,
    GridWriter,
    cell_id,
    sqlite_real,
)
from linexcel.formula_source import cell_source_fields
from linexcel.refs import a1, parse_ref, quote_sheet, split_sheet_prefix
from linexcel.structured_refs import formula_tokens

#: Schema 5 keeps literal input cells in per-sheet grids (``cell_grid``).
SCHEMA_VERSION = 5
#: The file name carries the schema: an older linexcel, which does not check
#: the schema, finds no index and asks for an import instead of reading a
#: grid index as if its literal cells were blank.
INDEX_FILENAME = "index-v5.sqlite"
#: Index files of earlier schemas, left in place until the project is
#: reimported.
LEGACY_INDEX_FILENAMES = ("index.sqlite",)
MAX_PAGE_SIZE = 200
#: The node kinds that stand for one worksheet cell.
_CELL_KINDS = "kind IN ('cell','input','array_result')"
_CELL_KIND_SET = frozenset({"cell", "input", "array_result"})
#: Positions after a sheet's cells: a cell's opaque position in one sheet is
#: ``row * _COLUMN_SPAN + col``; nodes added after the import come later.
_COLUMN_SPAN = 16385
_LATE = 1 << 50
MAX_NEIGHBORHOOD_DEPTH = 20
_NODE_KEYS = (
    "id",
    "kind",
    "sheet",
    "cell",
    "row",
    "column",
    "formula",
    "numberFormat",
    "sourceDataType",
    "context",
    "value",
    "cached_value",
    "cachedValueKind",
    "cached_comparison_value",
    "cache_status",
    "provenance",
    "dependencies",
    "diagnostics",
    "reference_bindings",
    "volatile",
    "sourceFormula",
    "canonicalFormula",
    "formulaProvenance",
    "sourceValueProvenance",
)
_SHORT_KEYS = {key: chr(65 + i) for i, key in enumerate(_NODE_KEYS)}
_LONG_KEYS = {short: key for key, short in _SHORT_KEYS.items()}
_XSTRING_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _decode_xstring(value: str) -> str:
    """Decode OOXML ST_Xstring once, preserving escaped literal spellings.

    For example ``_x005F_x0041_`` means the literal ``_x0041_``, not ``A``.
    UTF-16 surrogate pairs encode one Unicode character; an unpaired surrogate
    is invalid source data, not a replacement character or a readable value.
    """
    if "_x" not in value:
        return value
    decoded = _XSTRING_ESCAPE.sub(lambda match: chr(int(match[1], 16)), value)
    if any(0xD800 <= ord(character) <= 0xDFFF for character in decoded):
        decoded = decoded.encode("utf-16-le", "surrogatepass").decode("utf-16-le")
    return decoded


def _dump(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False)


def _pack(node):
    # Per-cell native compression and short field names avoid writing several
    # gigabytes of repeated schema, aliases and saved values for a large sheet.
    packed = {}
    for key, value in node.items():
        if key in {
            "label",
            "cachedValue",
            "valueSource",
            "formulaTokens",
            "patternId",
            "patternMemberCount",
        }:
            continue
        if (
            key in {"cached_value", "cached_comparison_value"}
            and value == node.get("value")
            and type(value) is type(node.get("value"))
        ):
            continue
        packed[_SHORT_KEYS.get(key, key)] = value
    return zlib.compress(_dump(packed).encode(), level=1)


def _unpack(data):
    packed = json.loads(zlib.decompress(data))
    node = {_LONG_KEYS.get(key, key): value for key, value in packed.items()}
    if "value" in node:
        node.setdefault("cached_value", node["value"])
        node.setdefault("cached_comparison_value", node["value"])
    if node.get("kind") == "array_result" and node.get("arrayAnchor"):
        # Members have no <f> of their own, but their stored values are array
        # caches, not literal inputs. Normalize older indexes on read as well.
        if node.get("provenance") == "saved_input":
            node["provenance"] = "saved_cache"
        if provenance := node.get("sourceValueProvenance"):
            provenance.update(role="formula_cache", arrayAnchor=node["arrayAnchor"])
    return node


class IndexSchemaError(ValueError):
    """The index was written by another linexcel version: reimport the workbook."""

    def __init__(self, found):
        self.found = found
        super().__init__(
            f"Index schema {found or 'unknown'} is not {SCHEMA_VERSION}: "
            "reimport the workbook."
        )


def _ascii_lower(text: str) -> str:
    # SQLite's LIKE folds ASCII letters only.
    return text.translate(_ASCII_LOWER)


_ASCII_LOWER = {c: c + 32 for c in range(65, 91)}


def _like_search(query: str):
    """``x LIKE '%query%'`` as SQLite evaluates it, for one Python string."""
    needle = _ascii_lower(query)
    return lambda text: text is not None and needle in _ascii_lower(text)


def _numeric_search_value(text: str) -> bool:
    # search_value GLOB '*[0-9]*' AND NOT GLOB '*[^0-9.eE+-]*'
    return any(c.isdigit() and c.isascii() for c in text) and all(
        c in "0123456789.eE+-" for c in text
    )


def _unknown_array_owner(node, *, persisted=False):
    # Keep the saved scalar/provenance, while refusing any chosen array owner.
    # Persist as an array result so source-cell readers still retain the value.
    node.update(
        kind="array_result" if persisted else "unresolved",
        reference=node["id"],
        dependencies=[],
        arrayOwnershipUnknown=True,
    )
    node.pop("arrayAnchor", None)
    if node.get("provenance") == "saved_input":
        node["provenance"] = "saved_cache"
    if provenance := node.get("sourceValueProvenance"):
        provenance.pop("arrayAnchor", None)
        provenance["role"] = "formula_cache"
    diagnostic = "Overlapping declared arrays have unknown member ownership."
    if diagnostic not in node.setdefault("diagnostics", []):
        node["diagnostics"].append(diagnostic)
    return node


def file_sha256(path: Path) -> str:
    with path.open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


def _part(part: str, target: str) -> str:
    return posixpath.normpath(
        target.lstrip("/")
        if target.startswith("/")
        else posixpath.join(posixpath.dirname(part), target)
    )


def _relationships(archive, part):
    name = posixpath.join(
        posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
    )
    if name not in archive.namelist():
        return {}
    return {
        r.get("Id"): (r.get("Type", ""), r.get("Target", ""), r.get("TargetMode"))
        for r in ET.fromstring(archive.read(name))
    }


def _rid(element):
    return next(
        (v for k, v in element.attrib.items() if k.rsplit("}", 1)[-1] == "id"), None
    )


def _stream_elements(archive, member: str, tag: str):
    """Yield complete records and detach them, including empty worksheet rows."""
    with archive.open(member) as source:
        stack = []
        for event, element in ET.iterparse(source, events=("start", "end")):
            if event == "start":
                stack.append(element)
                continue
            local = element.tag.rsplit("}", 1)[-1]
            if local == tag:
                yield element
                if len(stack) > 1:
                    stack[-2].remove(element)
                element.clear()
            elif local == "row":
                if len(stack) > 1:
                    stack[-2].remove(element)
                element.clear()
            stack.pop()


class WorkbookIndex(Mapping):
    """A read-through node mapping; iteration streams rows, never a workbook."""

    def __init__(self, path: Path, *, writable=False, check_schema=True):
        self.path = Path(path)
        uri_path = str(Path(path).resolve())
        if os.name == "nt":
            if not uri_path.startswith("\\\\?\\"):
                uri_path = (
                    "\\\\?\\UNC\\" + uri_path[2:]
                    if uri_path.startswith("\\\\")
                    else "\\\\?\\" + uri_path
                )
            # Path.as_uri interprets \\?\ as a host. Encode the original Windows
            # pathname instead: SQLite needs its extended prefix past MAX_PATH.
            readonly_uri = "file:" + quote(uri_path, safe="/:") + "?mode=ro"
        else:
            readonly_uri = Path(uri_path).as_uri() + "?mode=ro"
        self.db = sqlite3.connect(
            uri_path if writable else readonly_uri,
            uri=not writable,
        )
        self.db.execute("PRAGMA cache_size=-8192")
        self.db.execute("PRAGMA temp_store=FILE")
        if not writable:
            # One read transaction for the reader's lifetime: every statement
            # reads the same snapshot, and the file lock is taken once. In
            # autocommit each SELECT takes and releases it, about 270 µs a
            # lookup on Windows against 8 µs here (critical review F24).
            self.db.execute("BEGIN")
        if not writable and check_schema:
            try:
                found = self.db.execute(
                    "SELECT value FROM metadata WHERE key='schemaVersion'"
                ).fetchone()
            except sqlite3.DatabaseError:
                found = None
            if found is None or found[0] != str(SCHEMA_VERSION):
                self.db.close()
                raise IndexSchemaError(found[0] if found else None)
        self._meta: dict[str, Any] | None = None
        self._writer: GridWriter | None = None
        self._grid: GridReader | None = None
        self._layout: tuple[int, list[str]] | None = None
        self._selected_rects: dict[str, list[tuple[int, int, int, int]]] = {}

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

    def close(self):
        if hasattr(self, "_label_index"):
            self._label_index.close()
        self.db.close()

    def cell_labels(self, sheet, row, column):
        from linexcel.cell_labels import NON_TEXT_CELL, CellLabelIndex

        if not hasattr(self, "_label_index"):

            def literal(s, r, c):
                hit = self.db.execute(
                    "SELECT data FROM nodes WHERE sheet=? AND row=? AND col=? "
                    "AND kind IN ('input','cell','array_result') LIMIT 1",
                    (s, r, c),
                ).fetchone()
                node = _unpack(hit[0]) if hit else self.grid.get(s, r, c)
                if node:
                    if node["kind"] == "array_result":
                        return NON_TEXT_CELL
                    if not node.get("formula") and node.get("sourceDataType") in {
                        "s",
                        "str",
                        "inlineStr",
                    }:
                        return node.get("value")
                    if node.get("formula") or node.get("value") is not None:
                        return NON_TEXT_CELL
                return None

            names = []
            for (blob,) in self.db.execute("SELECT data FROM nodes WHERE kind='name'"):
                node = _unpack(blob)
                names.append(
                    dict(
                        name=node.get("name"),
                        expression=node.get("formula"),
                        scope="sheet" if node.get("sheet") else "workbook",
                        scope_sheet=node.get("sheet"),
                    )
                )
            self._label_index = CellLabelIndex(
                literal,
                tables=(
                    json.loads(r[0])
                    for r in self.db.execute("SELECT data FROM table_definitions")
                ),
                merges=(
                    (
                        (s, (r1, c1, r2, c2))
                        for s, r1, c1, r2, c2 in self.db.execute(
                            "SELECT * FROM merge_definitions"
                        )
                    )
                    if self.db.execute(
                        "SELECT 1 FROM sqlite_master WHERE type='table' "
                        "AND name='merge_definitions'"
                    ).fetchone()
                    else ()
                ),
                names=names,
            )
        return self._label_index.labels(sheet, row, column)

    @property
    def meta(self) -> dict[str, Any]:
        if self._meta is None:
            row = self.db.execute(
                "SELECT value FROM metadata WHERE key='meta'"
            ).fetchone()
            self._meta = json.loads(row[0]) if row else {}
        return self._meta

    # Grid cells (index schema 5). A literal input cell is either a node or a
    # grid cell, never both; every read below merges the two.

    @property
    def grid(self) -> GridReader:
        if self._grid is None:
            self._grid = GridReader(
                self.db, self._metadata_value("gridEpoch1904") == "1"
            )
        return self._grid

    def _metadata_value(self, key):
        row = self.db.execute(
            "SELECT value FROM metadata WHERE key=?", (key,)
        ).fetchone()
        return row[0] if row else None

    def _grid_cell(self, key):
        """``(sheet, row, col, raw)`` when ``key`` is a grid cell's id."""
        located = self.grid.locate(key)
        if located is None:
            return None
        raw = self.grid.raw(*located)
        return None if raw is None else (*located, raw)

    def _layout_info(self):
        """``(last seq stored by the import's reader, sheets in reading order)``."""
        if self._layout is None:
            end = self._metadata_value("gridImportEnd")
            order = json.loads(self._metadata_value("gridSheetOrder") or "[]")
            if end is None:
                end = self.db.execute("SELECT max(seq) FROM nodes").fetchone()[0]
            end = int(end or 0)
            known = set(order)
            for (sheet,) in self.db.execute(
                f"SELECT DISTINCT sheet FROM nodes WHERE {_CELL_KINDS} AND seq<=? "
                "AND sheet IS NOT NULL",
                (end,),
            ):
                if sheet not in known:
                    order.append(sheet)
                    known.add(sheet)
            for sheet in self.grid.sheets:
                if sheet not in known:
                    order.append(sheet)
                    known.add(sheet)
            self._layout = end, order
        return self._layout

    def _import_order(
        self, *, data=False, sheet=None, where="", args=(), grid=True, accept=None
    ):
        """Every node and grid cell in import order.

        Yields ``("node", row)``, ``row`` being ``(seq, id, kind, sheet, row,
        col[, data, pattern_id])``, or ``("grid", sheet, row, col, raw)``. The
        order is the import's: what the reader stored before the sheets
        (defined names), each sheet's cells row by row, then the nodes added
        afterwards (ranges, referenced blanks, external and VBA nodes). For
        worksheet XML that lists its cells in row order, as Excel writes it,
        this is the order the nodes were stored in. ``where``/``args`` filter
        node rows in SQL; ``grid`` says whether grid cells can match at all,
        and ``accept(sheet, row, col, raw)`` filters them.
        """
        columns = "seq,id,kind,sheet,row,col" + (",data,pattern_id" if data else "")
        end, order = self._layout_info()
        by_sheet, sheet_args = ("AND sheet IS ? ", (sheet,)) if sheet else ("", ())
        extra = f"AND ({where}) " if where else ""
        for row in self.db.execute(
            f"SELECT {columns} FROM nodes WHERE seq<=? AND NOT {_CELL_KINDS} "
            f"{by_sheet}{extra}ORDER BY seq",
            (end, *sheet_args, *args),
        ):
            yield "node", row
        for name in order:
            if sheet and name != sheet:
                continue
            nodes = (
                (row[4] or 0, row[5] or 0, 0, row)
                for row in self.db.execute(
                    f"SELECT {columns} FROM nodes WHERE sheet=? AND seq<=? "
                    f"AND {_CELL_KINDS} {extra}ORDER BY row,col,seq",
                    (name, end, *args),
                )
            )
            if not grid or name not in self.grid.sheets:
                for *_, row in nodes:
                    yield "node", row
                continue
            cells = (
                (r, c, 1, (name, r, c, raw))
                for r, c, *raw in self.grid.scan(name)
                if accept is None or accept(name, r, c, raw)
            )
            for *_, source, item in heapq.merge(nodes, cells, key=_merge_key):
                if source:
                    yield "grid", *item
                else:
                    yield "node", item
        for row in self.db.execute(
            f"SELECT {columns} FROM nodes WHERE seq>? {by_sheet}{extra}ORDER BY seq",
            (end, *sheet_args, *args),
        ):
            yield "node", row

    def _sheet_merge(self, sheet, bounds, node_sql, node_args, *, accept=None):
        """One sheet's node rows and grid cells, by row and column.

        ``node_sql`` selects ``row, col`` first and orders by them. Yields
        ``(row, col, node_row, None)`` or ``(row, col, None, raw)``.
        """
        nodes = (
            (r[0] or 0, r[1] or 0, 0, r, None)
            for r in self.db.execute(node_sql, node_args)
        )
        if sheet not in self.grid.sheets:
            for row, col, _, record, _ in nodes:
                yield row, col, record, None
            return
        cells = (
            (r, c, 1, None, raw)
            for r, c, *raw in self.grid.scan(sheet, bounds)
            if accept is None or accept(r, c)
        )
        for row, col, _, record, raw in heapq.merge(nodes, cells, key=_merge_key):
            yield row, col, record, raw

    def _grid_node(self, sheet, row, col, raw):
        return self.grid.node(sheet, row, col, *raw)

    def __len__(self):
        return (
            self.db.execute("SELECT count(*) FROM nodes").fetchone()[0]
            + self.grid.total
        )

    def __iter__(self):
        for source, *item in self._import_order():
            yield item[0][1] if source == "node" else cell_id(*item[:3])

    def __contains__(self, key):
        return (
            self.db.execute("SELECT 1 FROM nodes WHERE id=?", (key,)).fetchone()
            is not None
            or self._grid_cell(key) is not None
        )

    def __getitem__(self, key):
        row = self.db.execute(
            "SELECT data,pattern_id FROM nodes WHERE id=?", (key,)
        ).fetchone()
        if row is not None:
            return self._decode(row)
        found = self._grid_cell(key)
        if found is None:
            raise KeyError(key)
        return self._finish(self._grid_node(*found), None)

    def array_anchors(self, sheet, bounds, *, limit):
        """Bounded array-group provenance, without enumerating output cells."""
        if not self.db.execute(
            "SELECT 1 FROM sqlite_master WHERE name='array_definitions'"
        ).fetchone():
            return []
        r1, c1, r2, c2 = bounds
        return [
            f"{quote_sheet(sheet)}!{anchor}"
            for (anchor,) in self.db.execute(
                "SELECT anchor FROM array_definitions WHERE sheet=? "
                "AND r1<=? AND r2>=? AND c1<=? AND c2>=? ORDER BY anchor LIMIT ?",
                (sheet, r2, r1, c2, c1, limit),
            )
        ]

    def implicit_array_member(self, key):
        if not isinstance(key, str):
            return None
        rect = parse_ref(key)
        if not rect or not rect.sheet or rect.ncells != 1 or rect.to_a1() != key:
            return None
        anchors = self.array_anchors(
            rect.sheet, (rect.r1, rect.c1, rect.r2, rect.c2), limit=2
        )
        if not anchors or (len(anchors) == 1 and key in anchors):
            return None
        node = {
            "id": key,
            "kind": "array_result",
            "sheet": rect.sheet,
            "cell": a1(rect.r1, rect.c1),
            "row": rect.r1,
            "column": rect.c1,
            "formula": None,
            "value": None,
            "cache_status": "blank",
            "provenance": "saved_blank",
            "dependencies": anchors,
            "diagnostics": [],
        }
        if len(anchors) == 1:
            node["arrayAnchor"] = anchors[0]
        else:
            _unknown_array_owner(node)
        return node

    def resolve_array_ownership(self, node):
        if node.get("arrayOwnershipUnknown"):
            return _unknown_array_owner(node)
        blank = node.get("kind") == "input" and node.get("provenance") == "saved_blank"
        if blank or node.get("kind") == "array_result" or node.get("arrayBounds"):
            member = self.implicit_array_member(node["id"])
            if member and member.get("arrayOwnershipUnknown"):
                return _unknown_array_owner(node)
            if blank and member:
                node.update(member)
        return node

    def _decode(self, row, *, context=False):
        return self._finish(_unpack(row[0]), row[1], context=context)

    def _finish(self, node, pattern_id, *, context=False):
        node = self.resolve_array_ownership(node)
        node["label"] = (
            node.get("name") or node.get("cell") or node.get("reference") or node["id"]
        )
        node["cachedValue"] = node.get("cached_value")
        node["valueSource"] = node.get("provenance", "unknown")
        if pattern_id:
            node["patternId"] = pattern_id
            count = self.db.execute(
                "SELECT member_count FROM patterns WHERE id=?", (pattern_id,)
            ).fetchone()
            if count:
                node["patternMemberCount"] = count[0]
        if context and node.get("cell") and not node.get("external"):
            node["cellLabels"] = self.cell_labels(
                node["sheet"], node["row"], node["column"]
            )
            nearby = [
                dict(
                    cell=label["cell"],
                    text=label["text"],
                    position="left" if label["method"] == "left_text" else "above",
                    truncated=label["truncated"],
                )
                for label in node["cellLabels"]
                if label["method"] in {"left_text", "above_text"}
            ]
            node.setdefault("context", {})["nearbyLabels"] = nearby
        if context and node.get("formula"):
            try:
                tokens = formula_tokens(node["formula"])
                if len(tokens) <= 512:
                    node["formulaTokens"] = [
                        {"value": t.value, "type": t.type, "subtype": t.subtype}
                        for t in tokens
                    ]
            except Exception:
                pass  # The import diagnostics already describe the failure.
        return node

    def values(self, *, context=False):
        """Every node in import order; ``context`` adds labels and tokens."""
        for source, *item in self._import_order(data=True):
            if source == "node":
                yield self._decode(item[0][6:], context=context)
            else:
                yield self._finish(self._grid_node(*item), None, context=context)

    def items(self):
        for node in self.values():
            yield node["id"], node

    # Import -------------------------------------------------------------

    def begin_grid(self, *, epoch1904):
        """Keep the literal cells the reader stores next in per-sheet grids."""
        self._writer = GridWriter(self.db, epoch1904=epoch1904, demote=self._store)
        self.db.execute(
            "INSERT OR REPLACE INTO metadata VALUES ('gridEpoch1904',?)",
            ("1" if epoch1904 else "0",),
        )

    def note_array(self, sheet, r1, c1, r2, c2):
        """A declared array or data table, whose cells are never grid cells."""
        if self._writer is not None:
            self._writer.note_array(sheet, r1, c1, r2, c2)

    def finish_grid(self):
        """Write the grids and record where the reader's nodes end."""
        writer, self._writer = self._writer, None
        if writer is None:
            return
        writer.close()
        end = self.db.execute("SELECT coalesce(max(seq),0) FROM nodes").fetchone()[0]
        self.db.executemany(
            "INSERT OR REPLACE INTO metadata VALUES (?,?)",
            [("gridImportEnd", str(end)), ("gridSheetOrder", json.dumps(writer.order))],
        )
        self._grid = None
        self._layout = None

    def put(self, node, *, only_if_absent=False):
        writer = self._writer
        if writer is not None:
            if not only_if_absent and writer.offer(node):
                return
        elif (
            node.get("kind") in {"input", "array_result"}
            and node.get("sheet") in self.grid.sheets
            and type(node.get("row")) is int
            and type(node.get("column")) is int
        ):
            sheet, row, col = node["sheet"], node["row"], node["column"]
            raw = self.grid.raw(sheet, row, col)
            if raw is not None:
                if only_if_absent:
                    return
                self._demote_grid_cell(sheet, row, col, raw)
        self._store(node, only_if_absent=only_if_absent)

    def _demote_grid_cell(self, sheet, row, col, raw):
        """Turn a grid cell into a node, before a later writer replaces it."""
        node = self._grid_node(sheet, row, col, raw)
        info = self.grid.sheets[sheet]
        if col > GRID_COLUMNS:
            self.db.execute(
                "DELETE FROM grid_wide WHERE sheet_no=? AND row=? AND col=?",
                (info.no, row, col),
            )
        else:
            self.db.execute(
                f'UPDATE "{info.table}" SET c{col}=NULL WHERE row=?', (row,)
            )
        self.db.execute("UPDATE grid_sheets SET cells=cells-1 WHERE no=?", (info.no,))
        info.cells -= 1
        self._store(node)

    def _store(self, node, *, only_if_absent=False):
        sql = "INSERT OR IGNORE" if only_if_absent else "INSERT"
        suffix = (
            ""
            if only_if_absent
            else (
                " ON CONFLICT(id) DO UPDATE SET "
                "data=excluded.data,pattern_id=excluded.pattern_id,kind=excluded.kind"
            )
        )
        self.db.execute(
            sql
            + (
                " INTO nodes(id,kind,sheet,row,col,data,pattern_id,formula,search_val"
                "ue) VALUES (?,?,?,?,?,?,?,?,?)"
            )
            + suffix,
            (
                node["id"],
                node["kind"],
                node.get("sheet"),
                node.get("row"),
                node.get("column"),
                _pack(node),
                node.get("patternId"),
                node.get("formula"),
                str(node["value"]) if node.get("value") is not None else None,
            ),
        )

    def range_members(self, sheet, bounds, *, kind=None):
        r1, c1, r2, c2 = bounds
        if kind in {None, "input"}:
            clause = "kind='input'" if kind else _CELL_KINDS
            for row, col, record, _ in self._sheet_merge(
                sheet,
                (r1, c1, r2, c2),
                "SELECT row,col,id FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
                f"AND col BETWEEN ? AND ? AND {clause} ORDER BY row,col",
                (sheet, r1, r2, c1, c2),
            ):
                yield record[2] if record else cell_id(sheet, row, col)
            return
        clause = (
            "kind IN ('cell','array_result')"
            if kind == "calculated"
            else "kind=?"
            if kind
            else "kind IN ('cell','input','array_result')"
        )
        arguments = (sheet, r1, r2, c1, c2) + (
            (kind,) if kind and kind != "calculated" else ()
        )
        for row in self.db.execute(
            "SELECT id FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
            f"AND col BETWEEN ? AND ? AND {clause} ORDER BY row,col",
            arguments,
        ):
            yield row[0]

    # Reads for other modules. Callers outside this module never query the
    # ``nodes`` table themselves, so how cells are stored can change here alone.

    def cell_ids(self):
        """Every cell (formula, literal, array member) id, in import order."""
        for source, *item in self._import_order(where=_CELL_KINDS):
            yield item[0][1] if source == "node" else cell_id(*item[:3])

    def sheet_cell_ids(self, sheet):
        """The cell ids of one sheet, row by row."""
        for row, col, record, _ in self._sheet_merge(
            sheet,
            None,
            f"SELECT row,col,id FROM nodes WHERE sheet=? AND {_CELL_KINDS} "
            "ORDER BY row,col",
            (sheet,),
        ):
            yield record[2] if record else cell_id(sheet, row, col)

    def node_id_at(self, sheet, row, col):
        """The id of a node stored at one position, or ``None``."""
        hit = self.db.execute(
            "SELECT id FROM nodes WHERE sheet=? AND row=? AND col=?",
            (sheet, row, col),
        ).fetchone()
        if hit:
            return hit[0]
        if self.grid.raw(sheet, row, col) is not None:
            return cell_id(sheet, row, col)
        return None

    def count_cells(self, sheet, bounds):
        """How many cells are stored in a rectangle ``(r1, c1, r2, c2)``."""
        r1, c1, r2, c2 = bounds
        return self.db.execute(
            f"SELECT count(*) FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
            f"AND col BETWEEN ? AND ? AND {_CELL_KINDS}",
            (sheet, r1, r2, c1, c2),
        ).fetchone()[0] + self.grid.count(sheet, (r1, c1, r2, c2))

    def sheet_cells_after(self, sheet, after, limit):
        """``(position, id, row, col, decode)`` of one sheet's cells.

        In import order, from just after the opaque ``position`` a previous
        record gave (0 to start), at most ``limit`` records. ``decode()``
        returns the stored node.
        """
        end, _ = self._layout_info()
        produced = 0
        if after < _LATE:
            top, left = divmod(after, _COLUMN_SPAN)
            for row, col, record, raw in self._sheet_merge(
                sheet,
                (max(top, 1), 1, 1 << 31, 1 << 31),
                "SELECT row,col,id,data FROM nodes WHERE sheet IS ? AND seq<=? "
                f"AND {_CELL_KINDS} AND (row>? OR (row=? AND col>?)) "
                "ORDER BY row,col,seq",
                (sheet, end, top, top, left),
                accept=lambda r, c: (r, c) > (top, left),
            ):
                if produced >= limit:
                    return
                produced += 1
                position = row * _COLUMN_SPAN + col
                if record:
                    yield position, record[2], row, col, partial(_unpack, record[3])
                else:
                    yield (
                        position,
                        cell_id(sheet, row, col),
                        row,
                        col,
                        partial(self._grid_node, sheet, row, col, raw),
                    )
            after = _LATE
        for seq, node_id, row, col, data in self.db.execute(
            "SELECT seq,id,row,col,data FROM nodes NOT INDEXED "
            f"WHERE sheet IS ? AND seq>? AND {_CELL_KINDS} ORDER BY seq LIMIT ?",
            (sheet, max(end, after - _LATE), limit - produced),
        ):
            yield _LATE + seq, node_id, row, col, partial(_unpack, data)

    def row_nodes(self, sheet, row, columns):
        """``{column: node}`` for the nodes stored in some columns of a row."""
        found = {
            col: _unpack(data)
            for col, data in self.db.execute(
                "SELECT col,data FROM nodes WHERE sheet IS ? AND row=? "
                "AND col IN (" + ",".join("?" for _ in columns) + ")",
                (sheet, row, *columns),
            )
        }
        if sheet in self.grid.sheets:
            for col in columns:
                if col not in found:
                    node = self.grid.get(sheet, row, col)
                    if node is not None:
                        found[col] = node
        return found

    def numeric_value_ids(self, ranges, limit):
        """Ids of cells whose saved value is a number in one of ``ranges``.

        ``ranges`` are ``(low, high)`` pairs. Only values spelled with digits
        and number punctuation count: a text such as "12 rue" casts to 12 and
        must not be read as a number. Import order, at most ``limit`` ids.
        """
        clause = " OR ".join(
            "CAST(search_value AS REAL) BETWEEN ? AND ?" for _ in ranges
        )
        formats = self.grid.formats

        def accept(_sheet, _row, _col, raw):
            # The text form a node's search_value would hold, filtered as the
            # SQL above filters it: str(value), never a date's ISO text.
            value, fmt, _ = raw
            kind, _, date_kind, _ = formats[fmt]
            if kind == "b" or date_kind:
                return False
            if kind == "n":
                number = float(value)
            elif _numeric_search_value(value):
                number = sqlite_real(value)
            else:
                return False
            return any(low <= number <= high for low, high in ranges)

        ids = []
        if limit <= 0:
            return ids
        for source, *item in self._import_order(
            where=f"{_CELL_KINDS} AND search_value GLOB '*[0-9]*' "
            "AND search_value NOT GLOB '*[^0-9.eE+-]*' "
            f"AND ({clause})",
            args=[bound for pair in ranges for bound in pair],
            accept=accept,
        ):
            ids.append(item[0][1] if source == "node" else cell_id(*item[:3]))
            if len(ids) >= limit:
                break
        return ids

    def nodes_by_sheet_and_pattern(self):
        """``(position, id, kind, sheet, pattern)`` of every node.

        Grouped by sheet, then by formula pattern (``''`` for none, and for
        a node without a sheet), in import order inside each group. A
        position is a unique integer; a grid cell's is negative.
        """
        sheets = {
            sheet for (sheet,) in self.db.execute("SELECT DISTINCT sheet FROM nodes")
        }
        for name in sorted(sheets | set(self.grid.sheets), key=lambda n: n or ""):
            yield from self._unpatterned(name)
            yield from self.db.execute(
                "SELECT seq,id,kind,coalesce(sheet,''),pattern_id FROM nodes "
                "WHERE sheet IS ? AND pattern_id IS NOT NULL ORDER BY pattern_id,seq",
                (name,),
            )

    def _unpatterned(self, sheet):
        """Nodes of one sheet without a formula pattern, and its grid cells."""
        info = self.grid.sheets.get(sheet)
        if info is None:
            yield from self.db.execute(
                "SELECT seq,id,kind,coalesce(sheet,''),'' FROM nodes "
                "WHERE sheet IS ? AND pattern_id IS NULL ORDER BY seq",
                (sheet,),
            )
            return
        end, _ = self._layout_info()
        yield from self.db.execute(
            "SELECT seq,id,kind,sheet,'' FROM nodes WHERE sheet=? AND seq<=? "
            f"AND pattern_id IS NULL AND NOT {_CELL_KINDS} ORDER BY seq",
            (sheet, end),
        )
        for row, col, record, _ in self._sheet_merge(
            sheet,
            None,
            "SELECT row,col,seq,id,kind FROM nodes WHERE sheet=? AND seq<=? "
            f"AND pattern_id IS NULL AND {_CELL_KINDS} ORDER BY row,col,seq",
            (sheet, end),
        ):
            if record:
                yield record[2], record[3], record[4], sheet, ""
            else:
                position = -((info.no << 34) | (row << 14) | col)
                yield position, cell_id(sheet, row, col), "input", sheet, ""
        yield from self.db.execute(
            "SELECT seq,id,kind,sheet,'' FROM nodes WHERE sheet=? AND seq>? "
            "AND pattern_id IS NULL ORDER BY seq",
            (sheet, end),
        )

    # A derived database that attaches this index under ``schema`` aggregates
    # it in SQLite, without moving millions of rows through Python; these
    # return the SQL that reads the index's own tables for such a query.

    def attached_pattern_sizes_sql(self, schema):
        """SQL of ``(sheet, pattern_id, size)``, ``''`` for none, per group."""
        return (
            "SELECT sheet,pattern_id,sum(size) size FROM ("
            "SELECT coalesce(sheet,'') sheet,coalesce(pattern_id,'') pattern_id,"
            f"count(*) size FROM {schema}.nodes "
            "GROUP BY coalesce(sheet,''),coalesce(pattern_id,'') "
            f"UNION ALL SELECT name,'',cells FROM {schema}.grid_sheets "
            "WHERE cells>0) GROUP BY sheet,pattern_id"
        )

    def attached_edges_sql(self, schema):
        """SQL of ``(source, target)`` for every edge a lineage view keeps.

        Edges to an array member whose owner was chosen among overlapping
        declared arrays by an older import are left out
        (``_array_edge_filter``).
        """
        return (
            f"SELECT e.source source,e.target target FROM {schema}.edges e "
            f"JOIN {schema}.nodes t ON t.id=e.target "
            f"WHERE {_array_edge_filter(self, schema=schema)}"
        )

    def bracketed_formula_nodes(self):
        """``(id, node)`` of nodes whose formula holds a ``[``, in table order.

        External references (``[1]Sheet!A1``) are among them; structured
        references are too, so a caller still checks each formula.
        """
        for node_id, data in self.db.execute(
            "SELECT id,data FROM nodes WHERE formula IS NOT NULL "
            "AND instr(formula,'[')>0"
        ):
            yield node_id, _unpack(data)

    def range_nodes(self):
        """``(id, node)`` of every range node, in table order."""
        for node_id, data in self.db.execute(
            "SELECT id,data FROM nodes WHERE kind='range'"
        ):
            yield node_id, _unpack(data)

    def sheet_cells(self, sheet):
        """The stored nodes of one sheet's cells, row by row, as stored.

        Unlike ``self[id]`` the nodes are not resolved for array ownership.
        """
        for row, col, record, raw in self._sheet_merge(
            sheet,
            None,
            f"SELECT row,col,data FROM nodes WHERE sheet=? AND {_CELL_KINDS} "
            "ORDER BY row,col",
            (sheet,),
        ):
            yield (
                _unpack(record[2]) if record else self._grid_node(sheet, row, col, raw)
            )

    def cell_records(self, sheet, *, exclude=(), limit):
        """``(row, col, decode)`` of one sheet's cells, row by row.

        Cells inside any ``exclude`` rectangle ``(r1, c1, r2, c2)`` are left
        out; at most ``limit`` records. ``decode()`` returns the stored node.
        """
        clauses, args = ["sheet=?", "row IS NOT NULL", "col IS NOT NULL"], [sheet]
        for r1, c1, r2, c2 in exclude:
            clauses.append("NOT (row BETWEEN ? AND ? AND col BETWEEN ? AND ?)")
            args.extend((r1, r2, c1, c2))
        produced = 0
        for row, col, record, raw in self._sheet_merge(
            sheet,
            None,
            "SELECT row,col,data FROM nodes WHERE "
            + " AND ".join(clauses)
            + f" AND {_CELL_KINDS} ORDER BY row,col LIMIT ?",
            (*args, limit),
            accept=lambda r, c: (
                not any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in exclude)
            ),
        ):
            if produced >= limit:
                return
            produced += 1
            if record:
                yield row, col, partial(_unpack, record[2])
            else:
                yield row, col, partial(self._grid_node, sheet, row, col, raw)

    def window_records(self, sheet, bounds, max_record_bytes):
        """The stored records of a rectangle's cells, for a bounded reader.

        Yields each record as stored (zlib-compressed JSON, see ``_pack``) so
        the caller can bound its inflation before decoding it, which protects
        against a hostile file. A record over ``max_record_bytes`` is yielded
        as ``None`` without being read.
        """
        r1, c1, r2, c2 = bounds
        for row, col, record, raw in self._sheet_merge(
            sheet,
            (r1, c1, r2, c2),
            "SELECT row,col,CASE WHEN length(data)<=? THEN data END "
            "FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
            f"AND col BETWEEN ? AND ? AND {_CELL_KINDS} ORDER BY row,col",
            (max_record_bytes, sheet, r1, r2, c1, c2),
        ):
            if record:
                yield record[2]
            else:
                # Packed like a stored record, for the same bounded decoder.
                packed = _pack(self._grid_node(sheet, row, col, raw))
                yield packed if len(packed) <= max_record_bytes else None

    def node_inventory(self):
        """``(sheet, kind, count)`` for every sheet (``None`` for names)."""
        counts = {
            (sheet, kind): count
            for sheet, kind, count in self.db.execute(
                "SELECT sheet,kind,count(*) FROM nodes GROUP BY sheet,kind"
            )
        }
        for name, info in self.grid.sheets.items():
            if info.cells:
                counts[name, "input"] = counts.get((name, "input"), 0) + info.cells
        return sorted(
            ((sheet, kind, count) for (sheet, kind), count in counts.items()),
            key=lambda item: (item[0] is not None, item[0] or "", item[1]),
        )

    def name_records(self):
        """``(id, sheet, node)`` of every defined name, in import order."""
        for node_id, sheet, data in self.db.execute(
            "SELECT id,sheet,data FROM nodes WHERE kind='name' ORDER BY seq"
        ):
            yield node_id, sheet, _unpack(data)

    def sheet_records(self, sheet):
        """``(id, kind, row, col, decode)`` of one sheet's nodes, row by row.

        Names excluded. ``decode()`` returns the stored node; a caller that
        skips a record never pays for decoding it.
        """
        for row, col, record, raw in self._sheet_merge(
            sheet,
            None,
            "SELECT row,col,id,kind,data FROM nodes WHERE sheet IS ? "
            "AND kind!='name' ORDER BY row,col,seq",
            (sheet,),
        ):
            if record:
                yield (
                    record[2],
                    record[3],
                    record[0],
                    record[1],
                    partial(_unpack, record[4]),
                )
            else:
                yield (
                    cell_id(sheet, row, col),
                    "input",
                    row,
                    col,
                    partial(self._grid_node, sheet, row, col, raw),
                )

    def prepare_input_selection(self):
        # The selection is keyed by position, so it is read back in sheet order
        # by walking its own rows. Joining a bare id list to ``nodes`` in that
        # order made SQLite walk the whole position index instead: 1.6 s per
        # read for 75 inputs of a 2.7-million-node index (critical review F26).
        self.db.execute(
            "CREATE TEMP TABLE IF NOT EXISTS selected_inputs "
            "(sheet TEXT NOT NULL, row INTEGER NOT NULL, col INTEGER NOT NULL, "
            "id TEXT NOT NULL, PRIMARY KEY (sheet, row, col, id)) WITHOUT ROWID"
        )
        self.db.execute("DELETE FROM selected_inputs")
        # Grid cells of selected ranges stay rectangles: a range of a million
        # literals is never copied cell by cell. Rectangles are kept disjoint.
        self._selected_rects: dict[str, list[tuple[int, int, int, int]]] = {}

    def select_input(self, node_id):
        found = self._grid_cell(node_id)
        if found is not None:
            self.db.execute(
                "INSERT OR IGNORE INTO selected_inputs VALUES (?,?,?,?)",
                (*found[:3], node_id),
            )
            return
        self.db.execute(
            "INSERT OR IGNORE INTO selected_inputs SELECT coalesce(sheet,''),"
            "coalesce(row,0),coalesce(col,0),id FROM nodes WHERE id=?",
            (node_id,),
        )

    def select_range_inputs(self, sheet, bounds):
        r1, c1, r2, c2 = bounds
        self.db.execute(
            "INSERT OR IGNORE INTO selected_inputs SELECT sheet,row,col,id "
            "FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
            "AND col BETWEEN ? AND ? AND kind='input'",
            (sheet, r1, r2, c1, c2),
        )
        if sheet in self.grid.sheets:
            rects = self._selected_rects.setdefault(sheet, [])
            pieces = [(r1, c1, r2, c2)]
            for existing in rects:
                pieces = [
                    part for piece in pieces for part in _subtract(piece, existing)
                ]
            rects.extend(pieces)

    def _selection(self):
        """``(sheet, row, col, node data | None, raw | None)``, in sheet order.

        Selected nodes and grid cells, each position once, ordered by sheet,
        row and column like the selection table.
        """
        table = (
            (sheet, row, col, 0, data, node_id)
            for sheet, row, col, node_id, data in self.db.execute(
                "SELECT s.sheet,s.row,s.col,s.id,n.data FROM selected_inputs s "
                "LEFT JOIN nodes n ON n.id=s.id ORDER BY s.sheet,s.row,s.col,s.id"
            )
        )
        ranges = []
        for sheet in sorted(self._selected_rects):
            scans = [
                (
                    (sheet, r, c, 1, None, raw)
                    for r, c, *raw in self.grid.scan(sheet, rect)
                )
                for rect in self._selected_rects[sheet]
            ]
            ranges.append(heapq.merge(*scans, key=_merge_key4))
        previous = None
        for sheet, row, col, _, data, extra in heapq.merge(
            table, *ranges, key=_merge_key4
        ):
            position = (sheet, row, col)
            if position == previous and data is None:
                continue  # a grid cell selected alone and by a range
            previous = position
            if data is not None:
                yield sheet, row, col, data, None
            elif isinstance(extra, str):
                found = self._grid_cell(extra)
                if found is not None:
                    yield sheet, row, col, None, found[3]
            else:
                yield sheet, row, col, None, extra

    def selected_inputs(self):
        """Selected literal inputs as stored nodes, by sheet, row and column."""
        for sheet, row, col, data, raw in self._selection():
            yield (
                _unpack(data)
                if data is not None
                else self._grid_node(sheet, row, col, raw)
            )

    def selected_stored_inputs(self):
        """The selected inputs stored as nodes; grid cells never have diagnostics."""
        for (data,) in self.db.execute(
            "SELECT n.data FROM selected_inputs s CROSS JOIN nodes n ON n.id=s.id "
            "ORDER BY s.sheet,s.row,s.col,s.id"
        ):
            yield _unpack(data)

    def selected_input_cells(self):
        """What the engine needs of each selected input, in sheet order.

        A stored node as such; for a grid cell only its position, saved
        value (a date's serial) and storage type, without building a node.
        """
        formats = self.grid.formats
        for sheet, row, col, data, raw in self._selection():
            if data is not None:
                yield _unpack(data)
                continue
            value, fmt, _ = raw
            kind = formats[fmt][0]
            yield {
                "sheet": sheet,
                "row": row,
                "column": col,
                "cached_comparison_value": bool(value) if kind == "b" else value,
                "sourceDataType": kind,
            }

    def selected_input_count(self):
        total = self.db.execute("SELECT count(*) FROM selected_inputs").fetchone()[0]
        for sheet, rects in self._selected_rects.items():
            total += sum(self.grid.count(sheet, rect) for rect in rects)
            for row, col in self.db.execute(
                "SELECT row,col FROM selected_inputs s WHERE sheet=? "
                "AND NOT EXISTS (SELECT 1 FROM nodes n WHERE n.id=s.id)",
                (sheet,),
            ):
                if any(r1 <= row <= r2 and c1 <= col <= c2 for r1, c1, r2, c2 in rects):
                    total -= 1
        return total

    def selected_input_sample(self, limit):
        """The first ``limit`` selected input ids in text order."""
        found = [
            r[0]
            for r in self.db.execute(
                "SELECT id FROM selected_inputs ORDER BY id LIMIT ?", (limit,)
            )
        ]
        for sheet, rects in self._selected_rects.items():
            found.extend(self._first_grid_ids(sheet, rects, limit))
        return sorted(set(found))[:limit]

    def _first_grid_ids(self, sheet, rects, limit):
        """The ``limit`` smallest ids, as text, of grid cells in ``rects``.

        Ids of one sheet compare by column letters, then by row digits as
        text; every id of column ``A`` sorts before any of ``AA``, then ``B``.
        """
        from linexcel.refs import num_to_col

        info = self.grid.sheets[sheet]
        columns = sorted(
            {
                c
                for _, c1, _, c2 in rects
                for c in self.grid.columns_with_runs(info, c1, c2)
            },
            key=num_to_col,
        )
        found: list[str] = []
        for col in columns:
            if len(found) >= limit:
                break
            rows = []
            for r1, c1, r2, c2 in rects:
                if not c1 <= col <= c2:
                    continue
                if col > GRID_COLUMNS:
                    rows.extend(
                        r
                        for (r,) in self.db.execute(
                            "SELECT row FROM grid_wide WHERE sheet_no=? AND col=? "
                            "AND row BETWEEN ? AND ? AND value IS NOT NULL",
                            (info.no, col, r1, r2),
                        )
                    )
                elif col <= info.width:
                    rows.extend(
                        r
                        for (r,) in self.db.execute(
                            f'SELECT row FROM "{info.table}" WHERE row BETWEEN ? '
                            f"AND ? AND c{col} IS NOT NULL",
                            (r1, r2),
                        )
                    )
            letters = num_to_col(col)
            for row in heapq.nsmallest(limit - len(found), rows, key=str):
                found.append(f"{quote_sheet(sheet)}!{letters}{row}")
        return found

    def node(self, node_id):
        row = self.db.execute(
            "SELECT data,pattern_id FROM nodes WHERE id=?", (node_id,)
        ).fetchone()
        if row:
            return self._decode(row, context=True)
        found = self._grid_cell(node_id)
        if found is not None:
            return self._finish(self._grid_node(*found), None, context=True)
        member = self.implicit_array_member(node_id)
        return self._decode((_pack(member), None), context=True) if member else None


def _merge_key(item):
    return item[0], item[1], item[2]


def _merge_key4(item):
    return item[0], item[1], item[2], item[3]


def _subtract(rect, other):
    """``rect`` minus ``other``, as at most four disjoint rectangles."""
    r1, c1, r2, c2 = rect
    o1, d1, o2, d2 = other
    if o1 > r2 or o2 < r1 or d1 > c2 or d2 < c1:
        return [rect]
    pieces = []
    if r1 < o1:
        pieces.append((r1, c1, o1 - 1, c2))
    if o2 < r2:
        pieces.append((o2 + 1, c1, r2, c2))
    top, bottom = max(r1, o1), min(r2, o2)
    if c1 < d1:
        pieces.append((top, c1, bottom, d1 - 1))
    if d2 < c2:
        pieces.append((top, d2 + 1, bottom, c2))
    return pieces


def _schema(db):
    db.executescript(GRID_SCHEMA)
    db.executescript("""
        PRAGMA journal_mode=DELETE;
        PRAGMA synchronous=NORMAL;
        CREATE TABLE metadata(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE nodes(seq INTEGER PRIMARY KEY,id TEXT UNIQUE NOT NULL,
            kind TEXT NOT NULL,sheet TEXT,row INTEGER,col INTEGER,
            data BLOB NOT NULL,pattern_id TEXT,formula TEXT,search_value TEXT);
        CREATE INDEX nodes_position ON nodes(sheet,row,col);
        CREATE INDEX nodes_kind ON nodes(kind,seq);
        CREATE INDEX nodes_pattern ON nodes(pattern_id,row,col);
        CREATE TABLE edges(source TEXT NOT NULL,target TEXT NOT NULL,
            PRIMARY KEY(source,target)) WITHOUT ROWID;
        CREATE INDEX edges_target ON edges(target);
        CREATE TABLE patterns(id TEXT PRIMARY KEY,sheet TEXT NOT NULL,
            member_count INTEGER NOT NULL,data TEXT NOT NULL);
        CREATE TABLE strings(id INTEGER PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE shared(sheet TEXT NOT NULL,id TEXT NOT NULL,
            origin TEXT NOT NULL,formula TEXT NOT NULL,
            PRIMARY KEY(sheet,id)) WITHOUT ROWID;
        CREATE TABLE arrays(sheet TEXT NOT NULL,r1 INTEGER,c1 INTEGER,
            r2 INTEGER,c2 INTEGER);
        CREATE TABLE array_definitions(sheet TEXT NOT NULL,anchor TEXT NOT NULL,
            r1 INTEGER,c1 INTEGER,r2 INTEGER,c2 INTEGER,kind TEXT NOT NULL,
            PRIMARY KEY(sheet,anchor)) WITHOUT ROWID;
        CREATE TABLE table_definitions(name_key TEXT NOT NULL,sheet TEXT NOT NULL,
            r1 INTEGER,c1 INTEGER,r2 INTEGER,c2 INTEGER,
            part TEXT NOT NULL,data TEXT NOT NULL);
        CREATE INDEX tables_name ON table_definitions(name_key);
        CREATE INDEX tables_position ON table_definitions(sheet,r1,r2);
        CREATE TABLE merge_definitions(sheet TEXT NOT NULL,r1 INTEGER,c1 INTEGER,
            r2 INTEGER,c2 INTEGER);
    """)


def _read_source(path, store):
    """Index literals and formula text in one forward XML pass per sheet."""
    from linexcel.lazy import _INPUT_ERROR_KINDS, _id, _json, _name_literal

    db = store.db
    source_digest = file_sha256(path)
    names = {}
    links = {}
    diagnostics = []
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        from linexcel.calculation_context import iteration_settings

        calculation_settings = iteration_settings(workbook.find("{*}calcPr"))
        rels = _relationships(archive, "xl/workbook.xml")
        prop = workbook.find("{*}workbookPr")
        epoch1904 = prop is not None and prop.get("date1904", "0") in {"1", "true"}
        epoch = CALENDAR_MAC_1904 if epoch1904 else CALENDAR_WINDOWS_1900
        store.begin_grid(epoch1904=epoch1904)
        formats = ["General"]
        style_part = next(
            (
                _part("xl/workbook.xml", v[1])
                for v in rels.values()
                if v[0].endswith("/styles")
            ),
            "xl/styles.xml",
        )
        if style_part in archive.namelist():
            try:
                root = ET.fromstring(archive.read(style_part))
                custom = {
                    int(e.attrib["numFmtId"]): e.get("formatCode", "General")
                    for e in root.findall("{*}numFmts/{*}numFmt")
                }
                formats = [
                    custom.get(
                        int(e.get("numFmtId", "0")),
                        BUILTIN_FORMATS.get(int(e.get("numFmtId", "0")), "General"),
                    )
                    for e in root.findall("{*}cellXfs/{*}xf")
                ] or ["General"]
            except (
                ET.ParseError,
                ValueError,
                KeyError,
                zipfile.BadZipFile,
                OSError,
            ) as exc:
                # Only the default style stays known. A cell with any other
                # style index keeps an unknown value (see the per-cell check
                # below) instead of a guessed number, date or duration.
                formats = ["General"]
                diagnostics.append(
                    {
                        "part": style_part,
                        "message": "Styles part unreadable: number formats are "
                        "unknown, so saved values of styled cells are unknown. "
                        f"{type(exc).__name__}: {exc}",
                    }
                )
        elif any(v[0].endswith("/styles") for v in rels.values()):
            # Declared but absent: the same unknown values as an unreadable
            # part, and the workbook must say so (critical review F17).
            diagnostics.append(
                {
                    "part": style_part,
                    "message": "Styles part declared by the workbook is missing: "
                    "number formats are unknown, so saved values of styled cells "
                    "are unknown.",
                }
            )
        # Elapsed-format detection does not strip literals itself. A quoted
        # "[h]" label on a date must not turn its cache into a timedelta.
        format_kinds = [
            (is_date_format(f), is_timedelta_format(re.sub(r'"[^"]*"|\\.', "", f)))
            for f in formats
        ]
        string_part = next(
            (
                _part("xl/workbook.xml", v[1])
                for v in rels.values()
                if v[0].endswith("/sharedStrings")
            ),
            "xl/sharedStrings.xml",
        )
        strings_declared = any(v[0].endswith("/sharedStrings") for v in rels.values())
        strings_present = string_part in archive.namelist()
        missing_strings = 0
        if strings_present:
            try:
                for index, element in enumerate(
                    _stream_elements(archive, string_part, "si")
                ):
                    # Phonetic guides are not displayed cell text.
                    value = "".join(
                        e.text or ""
                        for e in element.findall("{*}t") + element.findall("{*}r/{*}t")
                    )
                    try:
                        value = _decode_xstring(value)
                    except UnicodeError as exc:
                        diagnostics.append(
                            {
                                "part": string_part,
                                "message": "Invalid Unicode shared string "
                                f"{index}: {exc}",
                            }
                        )
                        continue
                    db.execute("INSERT INTO strings VALUES (?,?)", (index, value))
                    if index % 10000 == 0:
                        db.commit()
            except (ET.ParseError, zipfile.BadZipFile, OSError) as exc:
                diagnostics.append(
                    {
                        "part": string_part,
                        "message": "Shared-string table is incomplete; missing strings "
                        "are unavailable, never empty values: "
                        f"{type(exc).__name__}: {exc}",
                    }
                )
        sheets: list[dict[str, Any]] = []
        for item in workbook.findall("{*}sheets/{*}sheet"):
            relation = rels.get(_rid(item))
            sheets.append(
                {
                    "name": item.get("name"),
                    "state": item.get("state", "visible"),
                    "part": _part("xl/workbook.xml", relation[1]) if relation else None,
                    "formulas": 0,
                    "inputs": 0,
                    "status": "complete",
                    "activeFilters": False,
                    "hiddenRows": [],
                    "diagnostics": [],
                }
            )
        for index, link in enumerate(
            workbook.findall("{*}externalReferences/{*}externalReference"), 1
        ):
            relation = rels.get(_rid(link))
            if relation:
                part = _part("xl/workbook.xml", relation[1])
                target = next(
                    (
                        v[1]
                        for v in _relationships(archive, part).values()
                        if v[0].endswith("/externalLinkPath")
                    ),
                    None,
                )
                if target:
                    links[str(index)] = PureWindowsPath(unquote(target)).name
        for definition in workbook.findall("{*}definedNames/{*}definedName"):
            local = definition.get("localSheetId")
            name = definition.get("name", "")
            if not name:
                diagnostics.append({"message": "Unnamed defined-name entry skipped."})
                continue
            scope = None
            if local is not None:
                position = int(local) if re.fullmatch(r"\d+", local.strip()) else -1
                if not 0 <= position < len(sheets):
                    # An unknown scope could bind the wrong name. Skip it so that
                    # formulas reading it stay unresolved and refused.
                    diagnostics.append(
                        {
                            "part": "xl/workbook.xml",
                            "message": f"Defined name {name!r} skipped: invalid "
                            f"localSheetId {local!r}. Formulas that use it are "
                            "not resolved.",
                        }
                    )
                    continue
                scope = sheets[position]["name"]
            nid = f"name:{scope or '*'}:{name}"
            names[(scope, name.casefold())] = nid
            node = {
                "id": nid,
                "kind": "name",
                "name": name,
                "sheet": scope,
                "formula": "=" + (definition.text or "").lstrip("="),
                "value": None,
                "cached_value": None,
                "cache_status": "unknown",
                "dependencies": [],
                "diagnostics": [],
            }
            literal = _name_literal(node)
            if literal is not None:
                _, value, kind = literal
                node.update(value=value, valueKind=kind, provenance="defined_constant")
            store.put(node)
        for sheet in sheets:
            row_number = 0
            column_number = 0
            previous_address = None
            row_hidden = False
            hidden_columns = []
            try:
                if not sheet["part"]:
                    raise ValueError("Worksheet relationship is missing")
                sheet_relationships = _relationships(archive, sheet["part"])
                with archive.open(sheet["part"]) as source:
                    stack = []
                    for event, element in ET.iterparse(source, events=("start", "end")):
                        tag = element.tag.rsplit("}", 1)[-1]
                        if event == "start":
                            stack.append(element)
                            if tag == "filterColumn":
                                sheet["activeFilters"] = True
                            elif tag == "row":
                                try:
                                    row_number = int(element.get("r", row_number + 1))
                                except (ValueError, TypeError):
                                    row_number += 1
                                    sheet["status"] = "partial"
                                    message = (
                                        "Invalid row coordinate; omitted cell "
                                        "coordinates may be ambiguous."
                                    )
                                    sheet["diagnostics"].append(message)
                                    diagnostics.append(
                                        {"sheet": sheet["name"], "message": message}
                                    )
                                column_number = 0
                                previous_address = None
                                row_hidden = element.get("hidden", "0") in {"1", "true"}
                                if row_hidden:
                                    sheet["hiddenRows"].append(row_number)
                            elif tag == "col" and element.get("hidden", "0") in {
                                "1",
                                "true",
                            }:
                                try:
                                    hidden_columns.append(
                                        (
                                            int(element.get("min")),
                                            int(element.get("max")),
                                        )
                                    )
                                except (ValueError, TypeError):
                                    diagnostics.append(
                                        {
                                            "sheet": sheet["name"],
                                            "message": "Invalid hidden-column "
                                            "metadata; visibility is unavailable "
                                            "for that interval.",
                                        }
                                    )
                            continue
                        if tag == "c":
                            address = element.get("r")
                            try:
                                if address is None:
                                    previous_column = (
                                        coordinate_to_tuple(previous_address)[1]
                                        if previous_address
                                        else 0
                                    )
                                    address = a1(row_number, previous_column + 1)
                                previous_address = address
                            except (ValueError, TypeError, KeyError):
                                address = None
                            # Excel often persists millions of formatting-only
                            # cells. They are absent inputs, so detach immediately
                            # without coordinate/style/date allocation or SQL.
                            if len(element) == 0 and element.get("t", "n") == "n":
                                stack[-2].remove(element)
                                element.clear()
                                stack.pop()
                                continue
                            try:
                                if not isinstance(address, str):
                                    raise ValueError("Cell coordinate is unavailable")
                                row, column_number = coordinate_to_tuple(address)
                                if not (
                                    1 <= row <= 1048576 and 1 <= column_number <= 16384
                                ):
                                    raise ValueError(
                                        "Cell coordinate is outside Excel bounds"
                                    )
                                style = int(element.get("s", "0"))
                                if not 0 <= style < len(formats):
                                    raise ValueError("Cell style is missing or invalid")
                            except (ValueError, TypeError, KeyError) as exc:
                                message = (
                                    f"Cell metadata unreadable at {address}: {exc}"
                                )
                                try:
                                    if not isinstance(address, str):
                                        raise ValueError(
                                            "Cell coordinate is unavailable"
                                        )
                                    row, column_number = coordinate_to_tuple(address)
                                    if not (
                                        1 <= row <= 1048576
                                        and 1 <= column_number <= 16384
                                    ):
                                        raise ValueError("Invalid coordinate")
                                except (ValueError, TypeError, KeyError):
                                    sheet["status"] = "partial"
                                    sheet["diagnostics"].append(message)
                                    diagnostics.append(
                                        {"sheet": sheet["name"], "message": message}
                                    )
                                else:
                                    assert isinstance(address, str)
                                    formula_element = element.find("{*}f")
                                    formula = (
                                        "=" + (formula_element.text or "")
                                        if formula_element is not None
                                        else None
                                    )
                                    store.put(
                                        {
                                            "id": _id(sheet["name"], address),
                                            "kind": "cell"
                                            if formula is not None
                                            else "input",
                                            "sheet": sheet["name"],
                                            "cell": address,
                                            "row": row,
                                            "column": column_number,
                                            "formula": formula,
                                            **cell_source_fields(
                                                element,
                                                workbook_sha256=source_digest,
                                                worksheet_part=sheet["part"],
                                                sheet=sheet["name"],
                                                cell=address,
                                                canonical_formula=formula,
                                                canonical_origin="unresolved",
                                                unreadable_value=True,
                                            ),
                                            "value": None,
                                            "cache_status": "unknown",
                                            "dependencies": [],
                                            "diagnostics": [message],
                                        }
                                    )
                                    sheet[
                                        "formulas" if formula is not None else "inputs"
                                    ] += 1
                                stack[-2].remove(element)
                                element.clear()
                                stack.pop()
                                continue
                            assert isinstance(address, str)
                            formula_element = element.find("{*}f")
                            formula = None
                            problems = []
                            shared_master = None
                            translation_status = None
                            translated_formula = None
                            canonical_origin = "xml"
                            if formula_element is not None:
                                formula = "=" + (formula_element.text or "")
                                form_type = formula_element.get("t")
                                if form_type == "shared":
                                    shared_id = formula_element.get("si", "")
                                    if formula_element.text:
                                        db.execute(
                                            "INSERT OR REPLACE INTO shared "
                                            "VALUES (?,?,?,?)",
                                            (
                                                sheet["name"],
                                                shared_id,
                                                address,
                                                formula,
                                            ),
                                        )
                                    else:
                                        shared = db.execute(
                                            "SELECT origin,formula FROM shared "
                                            "WHERE sheet=? AND id=?",
                                            (sheet["name"], shared_id),
                                        ).fetchone()
                                        if shared:
                                            shared_master = (shared[0], shared[1][1:])
                                            try:
                                                formula = Translator(
                                                    shared[1], origin=shared[0]
                                                ).translate_formula(address)
                                                translated_formula = formula
                                                translation_status = "translated"
                                                canonical_origin = "shared_translation"
                                            except Exception as exc:
                                                translation_status = "failed"
                                                canonical_origin = "unresolved"
                                                problems.append(
                                                    "Shared formula translation "
                                                    f"failed: {exc}"
                                                )
                                        else:
                                            translation_status = "missing_master"
                                            canonical_origin = "unresolved"
                                            problems.append(
                                                "Shared formula anchor is missing "
                                                "or occurs after this cell."
                                            )
                                elif form_type in {"array", "dataTable"}:
                                    rect = parse_ref(
                                        formula_element.get("ref", ""), sheet["name"]
                                    )
                                    if rect:
                                        db.execute(
                                            "INSERT INTO arrays VALUES (?,?,?,?,?)",
                                            (
                                                sheet["name"],
                                                rect.r1,
                                                rect.c1,
                                                rect.r2,
                                                rect.c2,
                                            ),
                                        )
                                        store.note_array(
                                            sheet["name"],
                                            rect.r1,
                                            rect.c1,
                                            rect.r2,
                                            rect.c2,
                                        )
                                        if (
                                            form_type == "array"
                                            and formula_element.text
                                            and (rect.r1, rect.c1)
                                            == (row, column_number)
                                        ):
                                            db.execute(
                                                "INSERT INTO array_definitions "
                                                "VALUES (?,?,?,?,?,?,?)",
                                                (
                                                    sheet["name"],
                                                    address,
                                                    rect.r1,
                                                    rect.c1,
                                                    rect.r2,
                                                    rect.c2,
                                                    "array",
                                                ),
                                            )
                                        else:
                                            problems.append(
                                                "Unsupported array/data-table metadata."
                                            )
                                    else:
                                        problems.append(
                                            "Array/data-table rectangle is unavailable."
                                        )
                                    if not formula_element.text:
                                        formula = "=<array/data-table>"
                                        canonical_origin = "unresolved"
                            cell_type = element.get("t", "n")
                            if cell_type not in {
                                "n",
                                "s",
                                "str",
                                "inlineStr",
                                "b",
                                "d",
                                "e",
                            }:
                                problems.append(
                                    f"Unsupported stored cell type: {cell_type}"
                                )
                            raw = element.findtext("{*}v")
                            if raw == "" and cell_type not in {"str", "inlineStr"}:
                                raw = None
                            number_format = (
                                formats[style]
                                if 0 <= style < len(formats)
                                else "General"
                            )
                            value = None
                            comparison = None
                            try:
                                if cell_type == "inlineStr":
                                    text_element = element.find("{*}is")
                                    value = (
                                        ""
                                        if text_element is None
                                        else "".join(
                                            e.text or ""
                                            for e in text_element.findall("{*}t")
                                            + text_element.findall("{*}r/{*}t")
                                        )
                                    )
                                    value = _decode_xstring(value)
                                elif cell_type == "s":
                                    if raw is None:
                                        raise ValueError(
                                            "Shared-string index is missing"
                                        )
                                    hit = db.execute(
                                        "SELECT value FROM strings WHERE id=?",
                                        (int(raw),),
                                    ).fetchone()
                                    if hit is None:
                                        if not strings_present:
                                            missing_strings += 1
                                        raise ValueError(f"Missing shared string {raw}")
                                    value = hit[0]
                                elif cell_type in {"str", "e"}:
                                    value = (
                                        raw
                                        if raw is not None
                                        else ""
                                        if cell_type == "str" and formula is None
                                        else None
                                    )
                                    if cell_type == "str" and value is not None:
                                        value = _decode_xstring(value)
                                elif cell_type == "b" and raw is not None:
                                    if raw not in {"0", "1"}:
                                        raise ValueError(
                                            "Boolean cell value must be 0 or 1"
                                        )
                                    value = raw == "1"
                                elif cell_type == "d" and raw is not None:
                                    value = from_ISO8601(raw)
                                    comparison = to_excel(value, epoch)
                                elif raw is not None:
                                    value = (
                                        float(raw)
                                        if any(t in raw for t in ".eE")
                                        else int(raw)
                                    )
                                    if isinstance(value, float) and not math.isfinite(
                                        value
                                    ):
                                        raise ValueError(
                                            "Non-finite stored numeric value"
                                        )
                                    comparison = value
                                    if (
                                        style < len(format_kinds)
                                        and format_kinds[style][0]
                                    ):
                                        if not (not epoch1904 and 60 <= value < 61):
                                            value = from_excel(
                                                value,
                                                epoch,
                                                timedelta=format_kinds[style][1],
                                            )
                                if (
                                    cell_type == "e"
                                    and formula is None
                                    and value not in _INPUT_ERROR_KINDS
                                ):
                                    problems.append(
                                        "Unknown Excel error input is not "
                                        "supported for recalculation."
                                    )
                            except Exception as exc:
                                problems.append(
                                    "Stored cell value is unreadable: "
                                    f"{type(exc).__name__}: {exc}"
                                )
                                value = None
                                comparison = None
                            if (
                                formula is not None
                                or value is not None
                                or cell_type in {"inlineStr", "s", "str"}
                                or problems
                            ):
                                node = {
                                    "id": _id(sheet["name"], address),
                                    "kind": "cell" if formula is not None else "input",
                                    "sheet": sheet["name"],
                                    "cell": address,
                                    "row": row,
                                    "column": column_number,
                                    "formula": formula,
                                    **cell_source_fields(
                                        element,
                                        workbook_sha256=source_digest,
                                        worksheet_part=sheet["part"],
                                        sheet=sheet["name"],
                                        cell=address,
                                        canonical_formula=formula,
                                        canonical_origin=canonical_origin,
                                        shared_master=shared_master,
                                        translation_status=translation_status,
                                        translated_formula=translated_formula,
                                        unreadable_value=any(
                                            p.startswith(
                                                "Stored cell value is unreadable:"
                                            )
                                            for p in problems
                                        ),
                                    ),
                                    "numberFormat": number_format,
                                    "sourceDataType": "f"
                                    if formula is not None
                                    else cell_type,
                                    "context": {
                                        "sheetState": sheet["state"],
                                        "rowHidden": row_hidden,
                                        "columnHidden": any(
                                            first <= column_number <= last
                                            for first, last in hidden_columns
                                        ),
                                    },
                                    "value": _json(value),
                                    "cached_value": _json(value),
                                    "cachedValueKind": "error"
                                    if cell_type == "e"
                                    else "scalar",
                                    "cached_comparison_value": comparison
                                    if comparison is not None
                                    else _json(value),
                                    "cache_status": "unknown"
                                    if formula is not None and value is None
                                    else "saved",
                                    "provenance": "saved_cache"
                                    if formula is not None
                                    else "saved_input",
                                    "dependencies": [],
                                    "diagnostics": problems,
                                }
                                store.put(node)
                                sheet[
                                    "formulas" if formula is not None else "inputs"
                                ] += 1
                                if (sheet["formulas"] + sheet["inputs"]) % 10000 == 0:
                                    db.commit()
                            stack[-2].remove(element)
                            element.clear()
                        elif tag == "mergeCell":
                            rect = parse_ref(element.get("ref", ""))
                            if rect:
                                db.execute(
                                    "INSERT INTO merge_definitions VALUES (?,?,?,?,?)",
                                    (sheet["name"], rect.r1, rect.c1, rect.r2, rect.c2),
                                )
                            stack[-2].remove(element)
                            element.clear()
                        elif tag == "tablePart":
                            from linexcel.structured_refs import read_table

                            part = None
                            try:
                                relation = sheet_relationships.get(_rid(element))
                                if (
                                    not relation
                                    or not relation[0].endswith("/table")
                                    or relation[2] == "External"
                                ):
                                    raise ValueError(
                                        "Table relationship is missing or external."
                                    )
                                part = _part(sheet["part"], relation[1])
                                table = read_table(
                                    archive, part, sheet["name"], source_digest
                                )
                                db.execute(
                                    "INSERT INTO table_definitions "
                                    "VALUES (?,?,?,?,?,?,?,?)",
                                    (
                                        table["name"].casefold(),
                                        sheet["name"],
                                        *table["bounds"],
                                        part,
                                        _dump(table),
                                    ),
                                )
                            except (
                                ValueError,
                                KeyError,
                                ET.ParseError,
                                OSError,
                                zipfile.BadZipFile,
                            ) as exc:
                                diagnostics.append(
                                    {
                                        "sheet": sheet["name"],
                                        "part": part,
                                        "feature": "structured_tables",
                                        "code": getattr(
                                            exc, "code", "table_metadata_invalid"
                                        ),
                                        "message": f"Table metadata unavailable: {exc}",
                                    }
                                )
                            stack[-2].remove(element)
                            element.clear()
                        elif tag == "row":
                            stack[-2].remove(element)
                            element.clear()
                        stack.pop()
            except (
                ET.ParseError,
                ValueError,
                KeyError,
                zipfile.BadZipFile,
                OSError,
            ) as exc:
                message = f"Worksheet import incomplete: {type(exc).__name__}: {exc}"
                sheet["status"] = "partial"
                sheet["diagnostics"].append(message)
                diagnostics.append({"sheet": sheet["name"], "message": message})
            sheet.pop("part", None)
            db.commit()
        if not strings_present and (strings_declared or missing_strings):
            # Without the part every text cell is unreadable, never empty. A
            # cell says so; the workbook must too, or the index reads as
            # complete (critical review F17).
            declared = " declared by the workbook" if strings_declared else ""
            diagnostics.append(
                {
                    "part": string_part,
                    "message": f"Shared-string part{declared} is missing: "
                    f"{missing_strings} text cell(s) are unreadable, never empty "
                    "values.",
                }
            )
    store.finish_grid()
    db.commit()
    return (
        {
            "sheets": [s["name"] for s in sheets],
            "sheetDetails": sheets,
            "dateSystem": "1904" if epoch1904 else "1900",
            "calculationSettings": calculation_settings,
            "diagnostics": diagnostics,
        },
        names,
        links,
    )


def _reference_feature(reference: str) -> str:
    """Capability registry feature for one reference that failed to resolve."""
    if reference.endswith("#"):
        return "arrays_and_spills"
    qualifier, _ = split_sheet_prefix(reference)
    if qualifier and "[" in qualifier:
        return "external_references"
    if qualifier and ":" in qualifier:
        return "three_dimensional_references"
    return "references"


_LOCAL_SCOPE = re.compile(r"_xlfn\.(?=(?:LET|LAMBDA)\s*\()", re.IGNORECASE)
_PARAMETER = re.compile(r"_xlpm\.", re.IGNORECASE)


def _local_names(formula: str, tokens) -> set[str]:
    """Upper-cased LET/LAMBDA names bound by one formula, without prefixes.

    Reuses the static ``analyze`` rule (``resolver._bound_names``). Excel saves
    ``LET(x,…)`` as ``_xlfn.LET(_xlpm.x,…)``; that rule only recognises the
    bare spelling, so both prefixes are removed before parsing. An expression
    the parser refuses binds nothing, which leaves its names unresolved and
    the calculation refused.
    """
    if not any(
        t.type == "FUNC"
        and t.subtype == "OPEN"
        and t.value[:-1].upper().removeprefix("_XLFN.") in {"LET", "LAMBDA"}
        for t in tokens
    ):
        return set()
    import formualizer as fz

    from linexcel.engine import is_too_deep
    from linexcel.resolver import _bound_names

    expression = _PARAMETER.sub("", _LOCAL_SCOPE.sub("", formula))
    if not expression.startswith("="):
        expression = "=" + expression
    if is_too_deep(expression):
        return set()
    try:
        return _bound_names(fz.parse(expression).to_dict())
    except Exception:
        return set()


def _index_dependencies(store, metadata, names, links, refs_dir):
    from linexcel.lazy import _DYNAMIC, _VOLATILE, _id, _pattern_signature

    db = store.db
    sheets = metadata["sheets"]
    lookup = {sheet.casefold(): sheet for sheet in sheets}
    partial = {s["name"] for s in metadata["sheetDetails"] if s["status"] != "complete"}
    from linexcel.external_index import ExternalSources

    with tempfile.TemporaryDirectory(
        prefix="linexcel-refs-", dir=store.path.parent
    ) as temporary:
        external_sources = ExternalSources(
            store, refs_dir, temporary, links, metadata["dateSystem"]
        )

        def dependency(text, sheet, source_node):
            if text.endswith("#"):
                anchor = parse_ref(text[:-1], default_sheet=sheet)
                declared = (
                    db.execute(
                        "SELECT sheet,r1,c1,r2,c2 FROM array_definitions "
                        "WHERE sheet=? AND anchor=?",
                        (
                            lookup.get((anchor.sheet or "").casefold()),
                            a1(anchor.r1, anchor.c1),
                        ),
                    ).fetchone()
                    if anchor is not None and anchor.ncells == 1
                    else None
                )
                if not declared:
                    raise ValueError(
                        "Spill reference requires a declared array anchor: " + text
                    )
                dep = dependency(text[:-1], sheet, source_node)
                from linexcel.refs import Rect

                # The source extent is verified against the native spill before
                # returning any result; this also works on engines without #.
                source_node.setdefault("resolvedReferences", {})[text] = Rect(
                    *declared
                ).to_a1()
                return dep
            external = external_sources.explicit(text, source_node)
            if external is not None:
                return external
            if source_node.get("external"):
                return external_sources.local(text, source_node)
            prefix, body = split_sheet_prefix(text)
            scope = lookup.get(prefix.casefold()) if prefix is not None else sheet
            name_id = None
            if prefix is None or scope is not None:
                name_id = names.get((scope, body.casefold())) or names.get(
                    (None, body.casefold())
                )
            if name_id:
                return name_id
            from linexcel.static_refs import sheet_span

            span = sheet_span(text, sheets)
            if span is not None:
                if text not in source_node.get("threeDimensionalArguments", []):
                    raise ValueError(
                        "3D references require a direct argument of a supported "
                        "aggregate function: " + text
                    )
                nid = f"sheet-span:{text}"
                members = [
                    dependency(rect.to_a1(), sheet, source_node) for rect in span
                ]
                source_node.setdefault("resolvedReferences", {})[text] = ",".join(
                    rect.to_a1() for rect in span
                )
                store.put(
                    {
                        "id": nid,
                        "kind": "sheet_span",
                        "sheet": None,
                        "reference": text,
                        "sheets": [rect.sheet for rect in span],
                        "cell_count": sum(rect.ncells for rect in span),
                        "dependencies": members,
                        "diagnostics": [],
                    },
                    only_if_absent=True,
                )
                for member in members:
                    db.execute(
                        "INSERT OR IGNORE INTO edges VALUES (?,?)", (member, nid)
                    )
                return nid
            from linexcel.structured_refs import resolve_table_reference

            table_selection = resolve_table_reference(
                store, text, sheet, source_node.get("row"), source_node.get("column")
            )
            if table_selection is not None:
                table = table_selection.table
                source_node.setdefault("resolvedReferences", {})[text] = (
                    table_selection.rectangle.to_a1()
                )
                required = source_node.setdefault("nativeTables", [])
                if table["sourcePart"] not in required:
                    required.append(table["sourcePart"])
                rect = table_selection.rectangle
            else:
                rect = parse_ref(text, default_sheet=sheet)
            if rect and rect.sheet:
                rect = replace(rect, sheet=lookup.get(rect.sheet.casefold()))
            if (
                rect
                and rect.sheet in sheets
                and (table_selection is not None or "[" not in text)
            ):
                nid = rect.to_a1()
                if nid in store:
                    return nid
                if rect.ncells == 1:
                    node = {
                        "id": nid,
                        "kind": "input",
                        "sheet": rect.sheet,
                        "cell": a1(rect.r1, rect.c1),
                        "row": rect.r1,
                        "column": rect.c1,
                        "formula": None,
                        "value": None,
                        "cached_value": None,
                        "cache_status": "blank",
                        "provenance": "saved_blank",
                        "dependencies": [],
                        "diagnostics": [],
                    }
                else:
                    assert rect.sheet is not None
                    array_anchors = [
                        _id(rect.sheet, row[0])
                        for row in db.execute(
                            "SELECT anchor FROM array_definitions WHERE sheet=? "
                            "AND r1<=? AND r2>=? AND c1<=? AND c2>=?",
                            (rect.sheet, rect.r2, rect.r1, rect.c2, rect.c1),
                        )
                    ]
                    node = {
                        "id": nid,
                        "kind": "range",
                        "sheet": rect.sheet,
                        "reference": text,
                        "bounds": [rect.r1, rect.c1, rect.r2, rect.c2],
                        "cell_count": rect.ncells,
                        "dependencies": array_anchors,
                        "diagnostics": [],
                    }
                if rect.sheet in partial:
                    node["diagnostics"].append(
                        "Source worksheet is incomplete; missing cells "
                        "cannot be treated as "
                        "blank."
                    )
                store.put(node, only_if_absent=True)
                for member in node["dependencies"]:
                    db.execute(
                        "INSERT OR IGNORE INTO edges VALUES (?,?)", (member, nid)
                    )
                return nid
            nid = f"unresolved:{sheet or '*'}:{text}"
            store.put(
                {
                    "id": nid,
                    "kind": "unresolved",
                    "sheet": sheet,
                    "reference": text,
                    "dependencies": [],
                    "diagnostics": [f"Unresolved dependency: {text}"],
                },
                only_if_absent=True,
            )
            return nid

        try:
            # New blank/range nodes have no formulas. A keyset cursor stays bounded
            # and remains valid across commits and node updates.
            last = 0
            while True:
                batch = db.execute(
                    (
                        "SELECT seq,data FROM nodes WHERE seq>? "
                        "AND kind IN ('cell','name') "
                        "ORDER BY seq LIMIT 512"
                    ),
                    (last,),
                ).fetchall()
                if not batch:
                    break
                for seq, data in batch:
                    last = seq
                    node = _unpack(data)
                    if node["kind"] == "cell":
                        region = db.execute(
                            (
                                "SELECT 1 FROM arrays WHERE sheet=? "
                                "AND ? BETWEEN r1 AND r2 AND ? "
                                "BETWEEN c1 AND c2 LIMIT 1"
                            ),
                            (node["sheet"], node["row"], node["column"]),
                        ).fetchone()
                        declared = db.execute(
                            "SELECT r1,c1,r2,c2 FROM array_definitions "
                            "WHERE sheet=? AND anchor=?",
                            (node["sheet"], node["cell"]),
                        ).fetchone()
                        if declared:
                            node["arrayBounds"] = list(declared)
                        elif region:
                            node["diagnostics"].append(
                                "Cell belongs to an unsupported array/data-table "
                                "result region."
                            )
                    try:
                        tokens = formula_tokens(node["formula"])
                        if any(
                            t.type == "FUNC"
                            and t.subtype == "OPEN"
                            and t.value[:-1].upper().removeprefix("_XLFN.")
                            in {"INDIRECT", "OFFSET"}
                            for t in tokens
                        ):
                            from linexcel.dynamic_refs import rewrite_dynamic_formula
                            from linexcel.lazy import _binding_expression

                            def resolve_dynamic(reference, current_sheet):
                                member = dependency(reference, current_sheet, node)
                                resolved = store[member]
                                expression = _binding_expression(resolved)
                                if expression is None or resolved.get("diagnostics"):
                                    raise ValueError(
                                        "Unresolved dynamic source reference: "
                                        + reference
                                    )
                                if member not in node["dependencies"]:
                                    node["dependencies"].append(member)
                                db.execute(
                                    "INSERT OR IGNORE INTO edges VALUES (?,?)",
                                    (member, node["id"]),
                                )
                                return resolved, expression

                            try:
                                if node.get("external") or node["kind"] == "name":
                                    raise ValueError(
                                        "Dynamic references in external sources or "
                                        "defined names require additional "
                                        "binding context."
                                    )
                                plan = rewrite_dynamic_formula(
                                    node["formula"],
                                    node["sheet"],
                                    node["row"],
                                    node["column"],
                                    resolve_dynamic,
                                )
                                node["evaluationFormula"] = plan.formula
                                node["dynamicReferences"] = plan.evidence
                                node["volatile"] = True
                                tokens += formula_tokens(plan.formula)
                            except ValueError as exc:
                                node["diagnostics"].append(str(exc))
                                node.setdefault("capabilityDiagnostics", []).append(
                                    {
                                        "feature": "dynamic_references",
                                        "code": "dynamic_reference_unresolved",
                                        "message": str(exc),
                                    }
                                )
                        from linexcel.static_refs import aggregate_arguments

                        eligible = aggregate_arguments(tokens)
                        spans = {}
                        for position, token in enumerate(tokens):
                            if token.type == "OPERAND" and token.subtype == "RANGE":
                                qualifier, _ = split_sheet_prefix(token.value)
                                if (
                                    qualifier
                                    and ":" in qualifier
                                    and "[" not in qualifier
                                ):
                                    spans[token.value] = (
                                        spans.get(token.value, True)
                                        and position in eligible
                                    )
                        if spans:
                            node["threeDimensionalArguments"] = [
                                value for value, valid in spans.items() if valid
                            ]
                        local = _local_names(node["formula"], tokens)
                        for token in tokens:
                            if (
                                local
                                and token.type == "OPERAND"
                                and token.subtype == "RANGE"
                                and split_sheet_prefix(token.value)[0] is None
                            ):
                                body = token.value
                                parameter = body.lower().startswith("_xlpm.")
                                bare = body[6:] if parameter else body
                                # A bare name that is also a defined name keeps
                                # that dependency: a superset never hides a
                                # source. Excel's _xlpm. marks a parameter.
                                if bare.upper() in local and (
                                    parameter
                                    or not (
                                        names.get((node.get("sheet"), bare.casefold()))
                                        or names.get((None, bare.casefold()))
                                    )
                                ):
                                    continue
                            if token.type == "FUNC" and token.subtype == "OPEN":
                                function = (
                                    token.value[:-1].upper().removeprefix("_XLFN.")
                                )
                                if function in _VOLATILE:
                                    node["volatile"] = True
                                if function in _DYNAMIC:
                                    node["diagnostics"].append(
                                        "Dynamic/workbook-sensitive dependency: "
                                        f"{function}"
                                    )
                            if token.type == "OPERAND" and token.subtype == "RANGE":
                                try:
                                    dep = dependency(
                                        token.value, node.get("sheet"), node
                                    )
                                except Exception as exc:
                                    # Each reference is extracted on its own: a
                                    # failure names that reference and keeps the
                                    # later ones. The diagnostic still refuses
                                    # calculation of this formula.
                                    from linexcel.structured_refs import (
                                        TableReferenceError,
                                    )

                                    if isinstance(exc, TableReferenceError):
                                        code, feature = exc.code, "structured_tables"
                                        message = str(exc)
                                    else:
                                        code = "reference_unresolved"
                                        feature = _reference_feature(token.value)
                                        message = (
                                            f"Reference {token.value} could not be "
                                            f"resolved: {exc}"
                                        )
                                    node.setdefault("capabilityDiagnostics", []).append(
                                        {
                                            "code": code,
                                            "feature": feature,
                                            "reference": token.value,
                                            "message": message,
                                        }
                                    )
                                    node["diagnostics"].append(message)
                                    continue
                                node.setdefault("reference_bindings", {})[
                                    token.value
                                ] = dep
                                if dep not in node["dependencies"]:
                                    node["dependencies"].append(dep)
                                db.execute(
                                    "INSERT OR IGNORE INTO edges VALUES (?,?)",
                                    (dep, node["id"]),
                                )
                    except Exception as exc:
                        node["diagnostics"].append(
                            f"Formula tokenization failed: {exc}"
                        )
                    if node["kind"] == "cell":
                        signature, tokens, reason = _pattern_signature(node, store)
                        identity = ["copied-r1c1-v1", node["sheet"], tokens]
                        if reason:
                            identity += ["isolated", node["id"], node["formula"]]
                        key = (
                            "formula-sha256:"
                            + hashlib.sha256(_dump(identity).encode()).hexdigest()
                        )
                        node["patternId"] = key
                        group = {
                            "id": key,
                            "sheet": node["sheet"],
                            "signature": signature,
                            "status": "isolated" if reason else "canonical",
                            "diagnostics": [reason] if reason else [],
                            "representativeId": node["id"],
                            "representativeFormula": node["formula"],
                        }
                        db.execute(
                            (
                                "INSERT INTO patterns VALUES (?,?,1,?) "
                                "ON CONFLICT(id) DO UPDATE SET "
                                "member_count=member_count+1"
                            ),
                            (key, node["sheet"], _dump(group)),
                        )
                    store.put(node)
                db.commit()
            # Arrays can include saved literal-looking cells. Mark them in the
            # index so target closures refuse cached array results consistently.
            for sheet, r1, c1, r2, c2 in db.execute("SELECT * FROM arrays"):
                for row in db.execute(
                    (
                        "SELECT data,pattern_id FROM nodes WHERE sheet=? "
                        "AND row BETWEEN ? "
                        "AND ? AND col BETWEEN ? AND ? AND kind='input'"
                    ),
                    (sheet, r1, r2, c1, c2),
                ):
                    node = store._decode(row)
                    definition = db.execute(
                        "SELECT anchor FROM array_definitions WHERE sheet=? "
                        "AND r1=? AND c1=? AND r2=? AND c2=?",
                        (sheet, r1, c1, r2, c2),
                    ).fetchone()
                    member = store.implicit_array_member(node["id"])
                    if member and member.get("arrayOwnershipUnknown"):
                        _unknown_array_owner(node, persisted=True)
                    elif definition:
                        anchor_id = _id(sheet, definition[0])
                        node.update(
                            kind="array_result",
                            arrayAnchor=anchor_id,
                            dependencies=[anchor_id],
                        )
                        db.execute(
                            "INSERT OR IGNORE INTO edges VALUES (?,?)",
                            (anchor_id, node["id"]),
                        )
                    else:
                        node["diagnostics"].append(
                            "Cell belongs to an unsupported array/data-table "
                            "result region."
                        )
                    store.put(node)
            db.commit()
        finally:
            external_sources.close()


def build_index(
    path: Path,
    index_path: Path,
    refs_dir: Path | None = None,
    *,
    filename: str | None = None,
) -> dict:
    "Publish a complete disk index atomically; preserve any previous index on failure."
    path, index_path = Path(path), Path(index_path)
    index_path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(
        prefix=index_path.name + ".", suffix=".building", dir=index_path.parent
    )
    os.close(handle)
    temp_path = Path(temporary)
    try:
        with WorkbookIndex(temp_path, writable=True) as store:
            _schema(store.db)
            from linexcel.macros import OLE_MAGIC, index_vba

            with path.open("rb") as source:
                macro_only = source.read(8).startswith(OLE_MAGIC) or (
                    (filename or path.name).lower().endswith(".xlam")
                )
            if macro_only:
                metadata: dict[str, Any] = {
                    "sheets": [],
                    "sheetDetails": [],
                    "diagnostics": [],
                    "macroOnly": True,
                }
            else:
                metadata, names, links = _read_source(path, store)
                _index_dependencies(store, metadata, names, links, refs_dir)
            from linexcel.link_facts import index_facts

            index_facts(path, store, metadata)
            metadata["vba"] = index_vba(store, path, refs_dir, filename=filename)
            metadata["diagnostics"].extend(
                {"phase": "VBA", "message": warning}
                for warning in metadata["vba"]["warnings"]
            )
            inventory = store.node_inventory()
            counts: dict[str, int] = {}
            for _, kind, count in sorted(inventory, key=lambda item: item[1]):
                counts[kind] = counts.get(kind, 0) + count
            per_sheet = {}
            for sheet, kind, count in inventory:
                per_sheet.setdefault(sheet or "defined names", {})[kind] = count
            pattern_count, copied, copied_cells = store.db.execute(
                "SELECT count(*),coalesce(sum(member_count>1),0),coalesce(sum(CASE "
                "WHEN member_count>1 THEN member_count ELSE 0 END),0) FROM patterns"
            ).fetchone()
            metadata.update(
                schemaVersion=SCHEMA_VERSION,
                mode="structural",
                evaluated=False,
                lazy=True,
                storage="sqlite",
                filename=filename or path.name,
                sha256=file_sha256(path),
                nodeCount=sum(counts.values()),
                nodeCounts=counts,
                perSheetCounts=per_sheet,
                sheetNodeCounts={s: sum(c.values()) for s, c in per_sheet.items()},
                formulaCount=counts.get("cell", 0),
                inputCount=counts.get("input", 0),
                patternCount=pattern_count,
                formulaPatternCount=pattern_count,
                copiedPatternCount=copied,
                copiedFormulaCells=copied_cells,
                edgeCount=store.db.execute("SELECT count(*) FROM edges").fetchone()[0],
                status="partial" if metadata["diagnostics"] else "complete",
                reference_files=sorted(
                    p.name for p in refs_dir.iterdir() if p.is_file()
                )
                if refs_dir and refs_dir.is_dir()
                else [],
                patternSemantics=(
                    "Sheet-local copied-formula signatures: A1 references "
                    "become relative "
                    "R1C1 offsets with absolute/mixed anchors preserved. Names retain "
                    "their scope, strings and operators retain their tokens. Formulas "
                    "with structural diagnostics or ambiguous references "
                    "remain isolated. Groups describe matching templates, "
                    "not equal values or proof of a "
                    "historical copy action. Member ranges tile actual cells without "
                    "filling gaps. No formula evaluation at import."
                ),
                contextSemantics=(
                    "Literal nearby text snippets, not inferred labels. "
                    "At most one text "
                    "cell within four cells left/above, each capped at 240 characters. "
                    "Formula and calculation inputs are never truncated by this "
                    "presentation limit."
                ),
                limits=[
                    *(
                        [
                            "Add-in source report: worksheet cells, values and formula "
                            "evaluation are not analyzed."
                        ]
                        if macro_only
                        else []
                    ),
                    "VBA relations are static source heuristics; "
                    "macros are never executed.",
                    (
                        "Local declared table references resolve to exact source "
                        "regions. Supplied external files support bounded source "
                        "formula/range/name closures. Dynamic references require "
                        "established source address expressions; unresolved "
                        "contexts are refused. See the versioned capability registry."
                    ),
                    (
                        "Range members are expanded sparsely only for "
                        "requested calculations."
                    ),
                    "A missing saved formula value is unknown, never zero.",
                    (
                        "Browser pages and graphs are bounded views "
                        "of the complete disk "
                        "index."
                    ),
                ],
            )
            from linexcel.capabilities import capability_registry

            metadata["capabilities"] = capability_registry()
            metadata["tableCount"] = store.db.execute(
                "SELECT count(*) FROM table_definitions"
            ).fetchone()[0]
            store.db.execute(
                "INSERT INTO metadata VALUES ('meta',?)", (_dump(metadata),)
            )
            store.db.execute(
                "INSERT OR REPLACE INTO metadata VALUES ('schemaVersion',?)",
                (str(SCHEMA_VERSION),),
            )
            store.db.commit()
        os.replace(temp_path, index_path)
    finally:
        temp_path.unlink(missing_ok=True)
    return graph_page(index_path)


def get_node(index_path: Path, node_id: str) -> dict | None:
    with WorkbookIndex(index_path) as store:
        return store.node(node_id)


def _filters(sheet=None, query=None, kind=None):
    conditions, values = [], []
    if sheet:
        conditions.append("sheet=?")
        values.append(sheet)
    if kind:
        conditions.append("kind=?")
        values.append(kind)
    if query:
        # Literal substring search, with SQL metacharacters escaped.
        conditions.append(
            "(id LIKE ? ESCAPE '\\' OR formula LIKE ? ESCAPE '\\' OR search_value "
            "LIKE ? ESCAPE '\\')"
        )
        pattern = (
            "%"
            + query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            + "%"
        )
        values += [pattern] * 3
    return " WHERE " + " AND ".join(conditions) if conditions else "", values


def _array_edge_filter(store, *, schema="main"):
    """Reject old chosen-owner edges using coordinates, never decoded cells.

    Ordinary formula targets skip the lookup. Array-result targets stop at the
    second containing declaration; the sheet key bounds the declaration scan.
    """
    if not store.db.execute(
        "SELECT 1 FROM sqlite_master WHERE name='array_definitions'"
    ).fetchone():
        return "1"
    return (
        "(t.kind<>'array_result' OR NOT EXISTS ("
        f"SELECT 1 FROM {schema}.array_definitions a WHERE a.sheet=t.sheet "
        "AND a.r1<=t.row AND a.r2>=t.row AND a.c1<=t.col AND a.c2>=t.col "
        "LIMIT 1 OFFSET 1))"
    )


def _edge_count(store):
    total = store.db.execute("SELECT count(*) FROM edges").fetchone()[0]
    condition = _array_edge_filter(store)
    if condition != "1":
        # Keep the ordinary-workbook count fast: only array-result targets use
        # the geometry check, via nodes_kind and edges_target, without blobs.
        total -= store.db.execute(
            "SELECT count(*) FROM nodes t JOIN edges e ON e.target=t.id "
            f"WHERE t.kind='array_result' AND NOT {condition}"
        ).fetchone()[0]
    return total


def _with_source_kinds(store, rows):
    """Edge rows ``(source, target, source kind, target kind)``.

    The source kind comes from a ``LEFT JOIN`` on ``nodes``: an input source
    that is a grid cell has none there. An edge whose source is neither a
    node nor a grid cell is left out, as an inner join left it out before.
    """
    for source, target, source_kind, target_kind in rows:
        if source_kind is None:
            if store._grid_cell(source) is None:
                continue
            source_kind = "input"
        yield source, target, source_kind, target_kind


def _edges_for_nodes(store, ids):
    from linexcel.macros import static_edge_record

    if not ids:
        return []
    placeholders = ",".join("?" for _ in ids)
    rows = store.db.execute(
        "SELECT e.source,e.target,s.kind,t.kind FROM edges e "
        "LEFT JOIN nodes s ON s.id=e.source JOIN nodes t ON t.id=e.target "
        f"WHERE e.target IN ({placeholders}) "
        f"AND e.source IN ({placeholders}) AND {_array_edge_filter(store)} "
        "ORDER BY e.source,e.target",
        ids + ids,
    )
    return [
        static_edge_record(s, t, sk, tk)
        for s, t, sk, tk in _with_source_kinds(store, rows)
    ]


def _pattern(store, row, member_limit=MAX_PAGE_SIZE):
    from linexcel.lazy import _member_ranges

    data, count = row
    group = json.loads(data)
    members = [
        {"id": n, "row": r, "column": c}
        for n, r, c in store.db.execute(
            "SELECT id,row,col FROM nodes WHERE pattern_id=? ORDER BY row,col LIMIT ?",
            (group["id"], member_limit),
        )
    ]
    group.update(
        memberCount=count,
        memberIds=[n["id"] for n in members],
        ranges=_member_ranges(members) if members else [],
        membersTruncated=count > len(members),
    )
    return group


def graph_page(
    index_path: Path, *, offset=0, limit=200, sheet=None, query=None, kind=None
) -> dict:
    from itertools import islice

    offset, limit = max(0, int(offset)), max(1, min(MAX_PAGE_SIZE, int(limit)))
    with WorkbookIndex(index_path) as store:
        # An exact Excel reference uses the unique id index, or the grid.
        # Substring search remains available when no exact node matches.
        exact = None
        if query:
            hit = store.db.execute(
                "SELECT data,pattern_id,kind,sheet FROM nodes WHERE id=?"
                + (" AND sheet=?" if sheet else ""),
                [query, sheet] if sheet else [query],
            ).fetchone()
            if hit is not None:
                exact = (hit[2], lambda: store._decode(hit[:2], context=True))
            else:
                found = store._grid_cell(query)
                if found is not None and (not sheet or found[0] == sheet):
                    exact = (
                        "input",
                        lambda: store._finish(
                            store._grid_node(*found), None, context=True
                        ),
                    )
        if exact is not None:
            matches = [exact[1]] if not kind or exact[0] == kind else []
            total = len(matches)
            nodes = [make() for make in matches[offset : offset + limit]]
        else:
            where, args = _filters(sheet, query, kind)
            total = store.db.execute("SELECT count(*) FROM nodes" + where, args)
            total = total.fetchone()[0]
            where, args = _filters(None, query, kind)
            condition = where.removeprefix(" WHERE ")
            accept = None
            if query:
                like = _like_search(query)
                search = store.grid.search_text

                def accept(name, row, col, raw):
                    return like(cell_id(name, row, col)) or like(search(raw))

            grid = kind in {None, "input"}
            if grid:
                names = [sheet] if sheet else list(store.grid.sheets)
                for name in names:
                    if accept is None:
                        total += store.grid.count(name)
                    else:
                        total += sum(
                            1
                            for r, c, *raw in store.grid.scan(name)
                            if accept(name, r, c, raw)
                        )
            nodes = [
                store._decode(item[0][6:], context=True)
                if source == "node"
                else store._finish(store._grid_node(*item), None, context=True)
                for source, *item in islice(
                    store._import_order(
                        data=True,
                        sheet=sheet,
                        where=condition,
                        args=args,
                        grid=grid,
                        accept=accept,
                    ),
                    offset,
                    offset + limit,
                )
            ]
        patterns = []
        for key in dict.fromkeys(n["patternId"] for n in nodes if n.get("patternId")):
            row = store.db.execute(
                "SELECT data,member_count FROM patterns WHERE id=?", (key,)
            ).fetchone()
            patterns.append(_pattern(store, row, member_limit=20))
        return {
            "lazy": True,
            "nodes": nodes,
            "edges": _edges_for_nodes(store, [n["id"] for n in nodes]),
            "formulaPatterns": patterns,
            "meta": store.meta,
            "pagination": {
                "offset": offset,
                "limit": limit,
                "total": total,
                "hasMore": offset + len(nodes) < total,
            },
        }


def structure_page(index_path: Path, *, part="nodes", offset=0, limit=200) -> dict:
    """Page the complete topology without decoding cells or starting an engine.

    Edges are paged independently: induced node-page edges would silently lose
    dependencies whose endpoints are on different pages. Ranges remain grouped
    as structural nodes; this does not expand their implicit cell membership.
    """
    if part not in {"nodes", "edges"}:
        raise ValueError("Unknown structure part")
    offset, limit = max(0, int(offset)), max(1, min(MAX_PAGE_SIZE, int(limit)))
    with WorkbookIndex(index_path) as store:
        total = len(store) if part == "nodes" else _edge_count(store)
        if part == "nodes":
            from itertools import islice

            items = []
            for source, *item in islice(store._import_order(), offset, offset + limit):
                if source == "node":
                    nid, kind, sheet = item[0][1:4]
                    items.append({"id": nid, "kind": kind, "sheet": sheet})
                else:
                    items.append(
                        {"id": cell_id(*item[:3]), "kind": "input", "sheet": item[0]}
                    )
        else:
            from linexcel.macros import static_edge_record

            items = [
                static_edge_record(source, target, sk, tk)
                for source, target, sk, tk in _with_source_kinds(
                    store,
                    store.db.execute(
                        "SELECT e.source,e.target,s.kind,t.kind FROM edges e "
                        "LEFT JOIN nodes s ON s.id=e.source "
                        "JOIN nodes t ON t.id=e.target "
                        f"WHERE {_array_edge_filter(store)} "
                        "ORDER BY e.source,e.target LIMIT ? OFFSET ?",
                        (limit, offset),
                    ),
                )
            ]
        return {
            part: items,
            "pagination": {
                "offset": offset,
                "total": total,
                "hasMore": offset + len(items) < total,
            },
        }


def patterns_page(index_path: Path, *, offset=0, limit=200, sheet=None) -> dict:
    offset, limit = max(0, int(offset)), max(1, min(MAX_PAGE_SIZE, int(limit)))
    with WorkbookIndex(index_path) as store:
        where, args = (" WHERE sheet=?", [sheet]) if sheet else ("", [])
        total = store.db.execute(
            "SELECT count(*) FROM patterns" + where, args
        ).fetchone()[0]
        groups = [
            _pattern(store, row, member_limit=20)
            for row in store.db.execute(
                "SELECT data,member_count FROM patterns"
                + where
                + (
                    " ORDER BY sheet,json_extract(data,'$.representativeId') LIMIT ? "
                    "OFFSET ?"
                ),
                args + [limit, offset],
            )
        ]
        return {
            "formulaPatterns": groups,
            "pagination": {
                "offset": offset,
                "limit": limit,
                "total": total,
                "hasMore": offset + len(groups) < total,
            },
        }


_MEMBER_KINDS = ("cell", "input", "array_result")


class _RangeMembership:
    """Range membership from indexed coordinates, without decoding cell blobs.

    A range node has no edge from its member cells in the ``edges`` table, so a
    cell read only through ``SUM(C2:C11)`` would otherwise look unused.
    ``lineage_ranges`` (written by ``link_facts.index_facts``) holds the local
    range bounds; an older index without it decodes its range nodes instead.
    """

    def __init__(self, store):
        self.store = store
        self.indexed = bool(
            store.db.execute(
                "SELECT 1 FROM sqlite_master WHERE name='lineage_ranges'"
            ).fetchone()
        )

    def _fallback_ranges(self, sheet):
        for nid, data in self.store.db.execute(
            "SELECT id,data FROM nodes WHERE kind='range' AND sheet=?", (sheet,)
        ):
            node = _unpack(data)
            if not node.get("external") and node.get("bounds"):
                yield nid, node["bounds"]

    def bounds(self, node_id, sheet):
        if self.indexed:
            row = self.store.db.execute(
                "SELECT r1,c1,r2,c2 FROM lineage_ranges WHERE node=?", (node_id,)
            ).fetchone()
            return list(row) if row else None
        return next(
            (b for nid, b in self._fallback_ranges(sheet) if nid == node_id), None
        )

    def containing(self, sheet, row, col, limit):
        """Local ranges that contain one cell, at most ``limit``."""
        if self.indexed:
            return [
                r[0]
                for r in self.store.db.execute(
                    "SELECT node FROM lineage_ranges WHERE sheet=? "
                    "AND ? BETWEEN r1 AND r2 AND ? BETWEEN c1 AND c2 "
                    "ORDER BY node LIMIT ?",
                    (sheet.lower(), row, col, limit),
                )
            ]
        return sorted(
            nid
            for nid, (r1, c1, r2, c2) in self._fallback_ranges(sheet)
            if r1 <= row <= r2 and c1 <= col <= c2
        )[:limit]

    def members(self, sheet, bounds, limit):
        """Indexed cells inside a range, in sheet order, at most ``limit``."""
        from itertools import islice

        return list(islice(self.store.range_members(sheet, bounds), limit))

    def edges(self, ids):
        """Membership relations (member -> range) among already selected nodes."""
        if not ids:
            return []
        marks = ",".join("?" for _ in ids)
        rows = self.store.db.execute(
            f"SELECT id,kind,sheet,row,col FROM nodes WHERE id IN ({marks})", ids
        ).fetchall()
        stored = {row[0] for row in rows}
        for nid in ids:
            if nid not in stored and (found := self.store._grid_cell(nid)):
                rows.append((nid, "input", *found[:3]))
        ranges = [
            (nid, sheet, b)
            for nid, kind, sheet, _, _ in rows
            if kind == "range" and (b := self.bounds(nid, sheet))
        ]
        cells = [r for r in rows if r[1] in _MEMBER_KINDS and r[3] is not None]
        return [
            {"source": cell, "target": rid, "kind": "range"}
            for rid, rsheet, (r1, c1, r2, c2) in ranges
            for cell, _, sheet, row, col in cells
            if sheet.lower() == rsheet.lower() and r1 <= row <= r2 and c1 <= col <= c2
        ]


def _adjacent(store, membership, node_id, condition, limit):
    """Direct sources and usages of one node, including range membership.

    One bounded query per relation family; each family returns at most
    ``limit`` rows, which keeps the caller's ``truncated`` flag exact.
    """
    found = {
        row[0]
        for row in store.db.execute(
            (
                "SELECT e.source AS neighbor FROM edges e "
                "JOIN nodes t ON t.id=e.target "
                f"WHERE e.target=? AND {condition} "
                "UNION SELECT e.target FROM edges e "
                "JOIN nodes t ON t.id=e.target "
                f"WHERE e.source=? AND {condition} "
                "ORDER BY neighbor LIMIT ?"
            ),
            (node_id, node_id, limit),
        )
    }
    position = store.db.execute(
        "SELECT kind,sheet,row,col FROM nodes WHERE id=?", (node_id,)
    ).fetchone()
    if position is None and (cell := store._grid_cell(node_id)):
        position = ("input", *cell[:3])
    if position:
        kind, sheet, row, col = position
        if kind == "range" and sheet:
            bounds = membership.bounds(node_id, sheet)
            if bounds:
                found.update(membership.members(sheet, bounds, limit))
        elif kind in _MEMBER_KINDS and sheet and row is not None:
            found.update(membership.containing(sheet, row, col, limit))
    return sorted(found)


def graph_neighborhood(index_path: Path, node_id: str, limit=200, depth=1) -> dict:
    """Read an undirected neighborhood with independent depth and node budgets.

    A deeper traversal never expands the HTTP node budget. ``truncated`` means
    an eligible node within the requested depth was omitted by that budget;
    neighbors beyond the requested depth are intentionally outside the view.
    Range membership counts as a relation: a cell read through a range reaches
    that range, and a range reaches its indexed member cells.
    """
    limit = max(1, min(MAX_PAGE_SIZE, int(limit)))
    depth = max(0, min(MAX_NEIGHBORHOOD_DEPTH, int(depth)))
    with WorkbookIndex(index_path) as store:
        if node_id not in store:
            raise ValueError(f"Unknown node: {node_id}")
        ids = {node_id: None}
        frontier = [node_id]
        truncated = False
        condition = _array_edge_filter(store)
        membership = _RangeMembership(store)
        for _ in range(depth):
            if not frontier:
                break
            following = []
            for current in frontier:
                rows = [
                    (neighbor,)
                    for neighbor in _adjacent(
                        store, membership, current, condition, limit + 1
                    )
                ]
                for row in rows:
                    if row[0] in ids:
                        continue
                    if len(ids) >= limit:
                        truncated = True
                        break
                    ids[row[0]] = None
                    following.append(row[0])
                if truncated:
                    break
            if truncated:
                break
            frontier = following
        nodes = [store.node(nid) for nid in ids]
        edges = sorted(
            _edges_for_nodes(store, list(ids)) + membership.edges(list(ids)),
            key=lambda e: (e["source"], e["target"], e["kind"]),
        )
        return {
            "lazy": True,
            "nodes": nodes,
            "edges": edges,
            "formulaPatterns": [
                _pattern(
                    store,
                    store.db.execute(
                        "SELECT data,member_count FROM patterns WHERE id=?", (key,)
                    ).fetchone(),
                    member_limit=20,
                )
                for key in dict.fromkeys(
                    n["patternId"] for n in nodes if n.get("patternId")
                )
            ],
            "meta": store.meta,
            "pagination": {
                "limit": limit,
                "returned": len(nodes),
                "truncated": truncated,
                "total": store.meta["nodeCount"],
            },
            "nodeId": node_id,
            "depth": depth,
        }


def evidence_graph(index_path: Path, node_id: str | None = None) -> dict:
    """Small deterministic evidence sample, with exact workbook-level counts."""
    from itertools import islice

    if node_id:
        with WorkbookIndex(index_path) as store:
            node = store.node(node_id)
            if node is None:
                raise ValueError(f"Unknown node: {node_id}")
            ids = list(dict.fromkeys([node_id, *node.get("dependencies", [])[:12]]))
            groups = []
            if node.get("patternId"):
                row = store.db.execute(
                    "SELECT data,member_count FROM patterns WHERE id=?",
                    (node["patternId"],),
                ).fetchone()
                if row:
                    groups.append(_pattern(store, row, member_limit=20))
            return {
                "lazy": True,
                "nodes": [store.node(nid) for nid in ids],
                "edges": _edges_for_nodes(store, ids),
                "formulaPatterns": groups,
                "meta": store.meta,
                "pagination": {
                    "returned": len(ids),
                    "total": len(set([node_id, *node.get("dependencies", [])])),
                    "truncated": len(node.get("dependencies", [])) > 12,
                },
            }
    with WorkbookIndex(index_path) as store:
        ids = []
        sheets = store.meta["sheets"]
        # Spread the bounded sample across the workbook, including the last sheet.
        selected = (
            [
                sheets[round(i * (len(sheets) - 1) / min(17, len(sheets) - 1))]
                for i in range(min(18, len(sheets)))
            ]
            if len(sheets) > 1
            else sheets
        )

        def import_ids(**selection):
            for source, *item in store._import_order(**selection):
                yield item[0][1] if source == "node" else cell_id(*item[:3])

        # Formulas first, in import order; then any other node.
        for sheet in selected:
            row = store.db.execute(
                "SELECT id FROM nodes WHERE sheet=? AND kind='cell' ORDER BY seq "
                "LIMIT 1",
                (sheet,),
            ).fetchone()
            first = row[0] if row else next(import_ids(sheet=sheet), None)
            if first:
                ids.append(first)
        cells = [
            row[0]
            for row in store.db.execute(
                "SELECT id FROM nodes WHERE kind='cell' ORDER BY seq LIMIT 36"
            )
        ]
        others = import_ids(where="kind!='cell'")
        for nid in [*cells, *islice(others, 36 - len(cells))]:
            if len(ids) >= 18:
                break
            if nid not in ids:
                ids.append(nid)
        return {
            "lazy": True,
            "nodes": [store.node(nid) for nid in ids],
            "edges": _edges_for_nodes(store, ids),
            "formulaPatterns": [],
            "meta": store.meta,
            "pagination": {
                "returned": len(ids),
                "total": store.meta["nodeCount"],
                "truncated": len(ids) < store.meta["nodeCount"],
            },
        }


def export_structure(index_path: Path) -> dict:
    """Explicit compatibility export. The web API must use bounded pages instead."""
    from linexcel.macros import static_edge_record

    with WorkbookIndex(index_path) as store:
        nodes = list(store.values(context=True))
        patterns = [
            _pattern(store, row, member_limit=row[1])
            for row in store.db.execute(
                "SELECT data,member_count FROM patterns ORDER BY "
                "sheet,json_extract(data,'$.representativeId')"
            )
        ]
        edges = [
            static_edge_record(s, t, sk, tk)
            for s, t, sk, tk in _with_source_kinds(
                store,
                store.db.execute(
                    "SELECT e.source,e.target,s.kind,t.kind FROM edges e "
                    "LEFT JOIN nodes s ON s.id=e.source JOIN nodes t ON t.id=e.target "
                    f"WHERE {_array_edge_filter(store)} "
                    "ORDER BY e.source,e.target"
                ),
            )
        ]
        return {
            "nodes": nodes,
            "formulaPatterns": patterns,
            "edges": edges,
            "meta": store.meta,
        }
