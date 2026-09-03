"""Power Query (Get & Transform) — the queries a workbook is fed by.

Data that arrives through a query has no formula behind it. The cells hold
values, the calculation engine sees a plain range, and the thing that filled
it — a CSV on a share, a database, another table of the same workbook — lives
in a part of the package no spreadsheet engine ever opens. Read it back and
the range stops being a dead end.

Where it sits in the file: a ``customXml`` item whose schema is
``http://schemas.microsoft.com/DataMashup`` holds a UTF-16 document whose one
element is base64. Decoded, that is a ``uint32`` version, a ``uint32`` length,
then a ZIP whose ``Formulas/Section1.m`` carries the M source of every query
in plain text. ``xl/connections.xml`` names the query each connection loads,
and ``xl/queryTables/*.xml`` ties that connection to the table on a sheet — so
the landing range is reachable too.

What is read out of the M is deliberately shallow: the sources a query names
and the queries it chains from. M is a full language and this is not an
interpreter; a source built at run time is invisible here exactly as a VBA
range built at run time is invisible to the VBA reader.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree

#: The custom XML schema Excel stamps on the mashup part.
DATA_MASHUP_SCHEMA = "http://schemas.microsoft.com/DataMashup"
#: Where the M source of every query sits inside the mashup package.
SECTION_PART = "Formulas/Section1.m"
#: Ceiling on the mashup package, so a corrupt length field cannot ask for a
#: gigabyte of memory on a file that only claims to hold one.
MAX_MASHUP_BYTES = 32 * 1024 * 1024
#: How many query sources one warning line names before it says "and more".
MAX_QUERY_SOURCES_SHOWN = 6

#: M functions that name a source outside the workbook, and what kind of thing
#: they read. Only the ones that take the target as their first string argument
#: are listed: ``Excel.Workbook(File.Contents("x.xlsx"))`` names its file
#: through ``File.Contents``, so matching the inner call is enough.
CONNECTORS = {
    "File.Contents": "file",
    "Folder.Files": "folder",
    "Folder.Contents": "folder",
    "Web.Contents": "web",
    "Web.Page": "web",
    "Web.BrowserContents": "web",
    "SharePoint.Contents": "web",
    "SharePoint.Files": "web",
    "SharePoint.Tables": "web",
    "AzureStorage.Blobs": "web",
    "AzureStorage.DataLake": "web",
    "Odbc.Query": "database",
    "Odbc.DataSource": "database",
}
#: Every connector Microsoft and its partners ship follows the same naming, so
#: a suffix rule covers the ones no list can enumerate — ``Sql.Database``,
#: ``Snowflake.Databases``, ``Salesforce.Data``, a partner's ``Foo.Feed``.
#: Deliberately narrow: transform namespaces (``Table``, ``Text``, ``List``)
#: have no member ending this way, so nothing of the pipeline is mistaken for
#: a source.
DATABASE_SUFFIXES = ("Database", "Databases", "Tables", "Cubes", "Feed", "DataSource")

#: A query reading a table or a named range of the workbook it lives in.
_CURRENT_WORKBOOK_RE = re.compile(
    r'Excel\.CurrentWorkbook\s*\(\s*\)\s*\{\s*\[\s*Name\s*=\s*"((?:[^"]|"")*)"'
)
#: ``Namespace.Member("first string argument"`` — the shape every connector
#: uses to name what it reads.
_CALL_RE = re.compile(r'\b([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)+)\s*\(\s*"((?:[^"]|"")*)"')
#: The head of a section member: optional metadata, optional ``shared``, then
#: the name — quoted when it holds spaces — and the ``=`` that opens its body.
_MEMBER_RE = re.compile(
    r'\s*(?:\[[^\]]*\]\s*)?(?:shared\s+)?(?:#"((?:[^"]|"")*)"|([A-Za-z_][\w.]*))\s*=\s*',
    re.DOTALL,
)
#: An identifier, quoted or bare, as M writes it.
_IDENT_RE = re.compile(r'#"((?:[^"]|"")*)"|\b([A-Za-z_]\w*)\b')
#: The same, immediately followed by a lone ``=``: a binding, not a reference.
_BOUND_RE = re.compile(r'(?:#"((?:[^"]|"")*)"|\b([A-Za-z_]\w*)\b)\s*=(?![=>])')


@dataclass(frozen=True)
class QuerySource:
    """One thing a query reads: a table, another query, a file, a server."""

    kind: str
    target: str
    function: str

    def as_dict(self) -> dict[str, str]:
        return {"kind": self.kind, "target": self.target, "function": self.function}


@dataclass(frozen=True)
class Destination:
    """Where a query lands: the sheet, the range, the table Excel created."""

    sheet: str | None
    ref: str | None
    table: str | None

    def as_dict(self) -> dict[str, str | None]:
        return {"sheet": self.sheet, "ref": self.ref, "table": self.table}


@dataclass
class Query:
    """One Power Query query: its M source, what it reads, where it lands."""

    name: str
    source: str
    sources: list[QuerySource] = field(default_factory=list)
    loaded_to: list[Destination] = field(default_factory=list)

    @property
    def loaded(self) -> bool:
        """``False`` for a connection-only query — computed, never written."""
        return bool(self.loaded_to)

    def outside_sources(self) -> list[QuerySource]:
        """The sources that are not in this workbook, so not in the graph."""
        return [s for s in self.sources if s.kind not in ("table", "query")]


def read_queries(data: bytes) -> list[Query]:
    """Every query of a workbook, in the order the mashup declares them.

    An empty list means what it says: no Power Query in the file, or a mashup
    part this reader could not make sense of. Either way the analysis carries
    on — a query that cannot be read costs the graph a node, not a run.
    """
    section = read_section(data)
    if not section:
        return []
    members = parse_section(section)
    if not members:
        return []
    destinations = read_destinations(data)
    names = set(members)
    return [
        Query(
            name=name,
            source=body,
            sources=scan_sources(body, names - {name}),
            loaded_to=destinations.get(name.casefold(), []),
        )
        for name, body in members.items()
    ]


def query_warning(queries: list[Query]) -> str | None:
    """One line for the queries that feed the workbook from outside it.

    A query whose source is a file, a URL or a server is a dependency of the
    same nature as a link to another workbook: the values it produced are in
    the file, what produced them is not. The graph shows the query and names
    its source; nobody should read that as the source having been checked.
    """
    if not queries:
        return None
    loaded = sum(1 for query in queries if query.loaded)
    if len(queries) == 1:
        head = "1 Power Query query feeds this workbook" + (
            ", loaded onto a sheet."
            if loaded
            else ", loaded nowhere (connection only)."
        )
    else:
        head = (
            f"{len(queries)} Power Query queries feed this workbook, "
            f"{loaded} of them loaded onto a sheet."
        )
    parts = [head]
    outside = sorted(
        {source.target for query in queries for source in query.outside_sources()}
    )
    if outside:
        shown = ", ".join(outside[:MAX_QUERY_SOURCES_SHOWN])
        if len(outside) > MAX_QUERY_SOURCES_SHOWN:
            shown += f", … (+{len(outside) - MAX_QUERY_SOURCES_SHOWN})"
        parts.append(
            f"Their data comes from outside the file and was not read: {shown}."
        )
    return " ".join(parts)


def read_section(data: bytes) -> str | None:
    """The ``Section1.m`` text of the workbook's mashup, if it has one."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for part in zf.namelist():
                if not re.fullmatch(r"customXml/item\d+\.xml", part):
                    continue
                root = _mashup_root(zf.read(part))
                if root is None:
                    continue
                section = _unpack_mashup(root.text or "")
                if section is not None:
                    return section
    except Exception:
        return None
    return None


def parse_section(text: str) -> dict[str, str]:
    """``section Section1; shared X = …;`` → ``{name: M source}``.

    >>> parse_section('section Section1;\\nshared Sales = let x = 1 in x;')
    {'Sales': 'let x = 1 in x'}
    """
    out: dict[str, str] = {}
    for statement in _statements(text):
        head = _MEMBER_RE.match(statement)
        if head is None:
            continue
        name = _identifier(head.group(1), head.group(2))
        body = statement[head.end() :].strip()
        if name and body:
            out[name] = body
    return out


def scan_sources(m_source: str, other_queries: set[str]) -> list[QuerySource]:
    """What one query reads, read off its M source.

    Three kinds of answer: a table or named range of this workbook, another
    query of the same file, or something outside both — a path, a URL, a
    server. Anything an expression computes rather than spells out is not
    here, and no static reader can put it there.
    """
    found: list[QuerySource] = []
    for match in _CURRENT_WORKBOOK_RE.finditer(m_source):
        found.append(
            QuerySource("table", _unquote(match.group(1)), "Excel.CurrentWorkbook")
        )
    for match in _CALL_RE.finditer(m_source):
        function, target = match.group(1), _unquote(match.group(2))
        kind = CONNECTORS.get(function)
        if kind is None and function.rpartition(".")[2].endswith(DATABASE_SUFFIXES):
            kind = "database"
        if kind is not None and target:
            found.append(QuerySource(kind, target, function))
    for name in _referenced_queries(m_source, other_queries):
        found.append(QuerySource("query", name, "let"))
    seen: set[tuple[str, str]] = set()
    unique: list[QuerySource] = []
    for source in found:
        key = (source.kind, source.target)
        if key not in seen:
            seen.add(key)
            unique.append(source)
    return unique


def read_destinations(data: bytes) -> dict[str, list[Destination]]:
    """``{query name, casefolded: [where it loads]}``.

    Follows the chain the format itself uses: worksheet → table part →
    query table → connection → the query the connection names. A query that
    loads nowhere — connection only, or straight into the data model — has no
    entry, which is the difference the report shows.
    """
    out: dict[str, list[Destination]] = defaultdict(list)
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            connections = _connection_queries(zf)
            if not connections:
                return {}
            book = _xml(zf, "xl/workbook.xml")
            if book is None:
                return {}
            rels = _rels(zf, "xl/workbook.xml")
            for element in book.iter():
                if _tag(element) != "sheet":
                    continue
                sheet = element.get("name")
                rid = next(
                    (v for k, v in element.attrib.items() if _tag_name(k) == "id"), None
                )
                part = rels.get(rid or "")
                if not sheet or not part:
                    continue
                for table_part in _rels(zf, part, kind="table").values():
                    table = _xml(zf, table_part)
                    if table is None:
                        continue
                    for qt_part in _rels(zf, table_part, kind="queryTable").values():
                        query_table = _xml(zf, qt_part)
                        if query_table is None:
                            continue
                        query = connections.get(query_table.get("connectionId") or "")
                        if not query:
                            continue
                        out[query.casefold()].append(
                            Destination(
                                sheet=sheet,
                                ref=table.get("ref"),
                                table=table.get("displayName") or table.get("name"),
                            )
                        )
    except Exception:
        return {}
    return dict(out)


# ---------------------------------------------------------------------------
# The mashup package
# ---------------------------------------------------------------------------


def _mashup_root(raw: bytes) -> Any:
    """The ``<DataMashup>`` element of a custom XML part, or ``None``."""
    for candidate in (raw, raw.decode("utf-16", "ignore")):
        try:
            root = ElementTree.fromstring(candidate)  # type: ignore[arg-type]
        except (ElementTree.ParseError, ValueError, UnicodeDecodeError):
            continue
        if root.tag == f"{{{DATA_MASHUP_SCHEMA}}}DataMashup":
            return root
        return None
    return None


def _unpack_mashup(encoded: str) -> str | None:
    """base64 → ``version | length | ZIP`` → the M source it holds."""
    try:
        blob = base64.b64decode(encoded, validate=False)
    except Exception:
        return None
    if len(blob) < 8:
        return None
    _version, length = struct.unpack_from("<II", blob, 0)
    if not 0 < length <= min(len(blob) - 8, MAX_MASHUP_BYTES):
        return None
    try:
        with zipfile.ZipFile(io.BytesIO(blob[8 : 8 + length])) as package:
            return package.read(SECTION_PART).decode("utf-8-sig")
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Reading M without interpreting it
# ---------------------------------------------------------------------------


def _statements(text: str) -> list[str]:
    """Split on the ``;`` that end section members, and only on those.

    A scanner rather than a ``split``: M holds strings, comments and quoted
    identifiers, any of which may carry a semicolon that ends nothing.
    """
    out: list[str] = []
    start = depth = index = 0
    end = len(text)
    while index < end:
        char = text[index]
        if char == '"':
            index = _skip_string(text, index)
            continue
        if text.startswith("//", index):
            newline = text.find("\n", index)
            index = end if newline < 0 else newline + 1
            continue
        if text.startswith("/*", index):
            close = text.find("*/", index + 2)
            index = end if close < 0 else close + 2
            continue
        if char in "([{":
            depth += 1
        elif char in ")]}":
            depth = max(0, depth - 1)
        elif char == ";" and depth == 0:
            out.append(text[start:index])
            start = index + 1
        index += 1
    if text[start:].strip():
        out.append(text[start:])
    return out


def _skip_string(text: str, index: int) -> int:
    """Index just past the string literal opening at ``index``."""
    index += 1
    end = len(text)
    while index < end:
        if text[index] == '"':
            if index + 1 < end and text[index + 1] == '"':
                index += 2
                continue
            return index + 1
        index += 1
    return end


def _scrub(text: str) -> str:
    """The M source with strings and comments blanked, identifiers kept.

    Quoted identifiers survive: ``#"Sales 2026"`` names a query, while
    ``"Sales 2026"`` is a value that happens to read the same.
    """
    out: list[str] = []
    index = 0
    end = len(text)
    while index < end:
        char = text[index]
        if char == '"':
            close = _skip_string(text, index)
            keep = index > 0 and text[index - 1] == "#"
            out.append(text[index:close] if keep else " " * (close - index))
            index = close
            continue
        if text.startswith("//", index):
            newline = text.find("\n", index)
            close = end if newline < 0 else newline
            out.append(" " * (close - index))
            index = close
            continue
        if text.startswith("/*", index):
            marker = text.find("*/", index + 2)
            close = end if marker < 0 else marker + 2
            out.append(" " * (close - index))
            index = close
            continue
        out.append(char)
        index += 1
    return "".join(out)


def _referenced_queries(m_source: str, others: set[str]) -> list[str]:
    """Queries this one chains from, named as M names them.

    A query is referenced by writing its name. The same name may also be a
    step of the local ``let``, and a step wins: it is what the expression
    actually reads. Everything bound in the body is therefore excluded, which
    is a heuristic, and one that errs towards claiming no edge rather than a
    wrong one.
    """
    if not others:
        return []
    body = _scrub(m_source)
    by_fold = {name.casefold(): name for name in others}
    bound = {
        _identifier(m.group(1), m.group(2)).casefold()
        for m in _BOUND_RE.finditer(body)
        if _identifier(m.group(1), m.group(2))
    }
    found: list[str] = []
    for match in _IDENT_RE.finditer(body):
        name = _identifier(match.group(1), match.group(2))
        fold = name.casefold()
        if fold in bound or fold not in by_fold:
            continue
        # ``Table.Sort`` is one name in two halves, not a query called Table.
        if match.group(2) is not None and _is_dotted(body, match.start(), match.end()):
            continue
        if by_fold[fold] not in found:
            found.append(by_fold[fold])
    return found


def _is_dotted(text: str, start: int, end: int) -> bool:
    """Is this bare identifier one half of a ``Namespace.Member`` name?"""
    before = text[start - 1] if start else ""
    after = text[end] if end < len(text) else ""
    return before == "." or after == "."


def _identifier(quoted: str | None, bare: str | None) -> str:
    """The name of an M identifier, whichever of the two spellings it uses."""
    if quoted is not None:
        return _unquote(quoted)
    return bare or ""


def _unquote(text: str) -> str:
    """``""`` is how M escapes a quote inside a quoted token."""
    return text.replace('""', '"')


# ---------------------------------------------------------------------------
# Package plumbing
# ---------------------------------------------------------------------------


def _connection_queries(zf: zipfile.ZipFile) -> dict[str, str]:
    """``{connection id: query name}`` for the mashup connections only."""
    root = _xml(zf, "xl/connections.xml")
    if root is None:
        return {}
    out: dict[str, str] = {}
    for connection in root:
        if _tag(connection) != "connection":
            continue
        cid = connection.get("id")
        if not cid:
            continue
        for child in connection:
            if _tag(child) != "dbPr":
                continue
            name = _mashup_location(child.get("connection") or "")
            if name is None:
                name = _mashup_command(child.get("command") or "")
            if name:
                out[cid] = name
    return out


def _mashup_location(connection: str) -> str | None:
    """``Provider=Microsoft.Mashup.OleDb.1;…;Location=Sales`` → ``Sales``."""
    if "Mashup" not in connection:
        return None
    for part in connection.split(";"):
        key, sep, value = part.partition("=")
        if sep and key.strip().casefold() == "location":
            return value.strip().strip('"')
    return None


def _mashup_command(command: str) -> str | None:
    """``SELECT * FROM [Sales]`` — the fallback when no Location is given."""
    match = re.search(r"FROM\s*\[([^\]]+)\]", command, re.IGNORECASE)
    return match.group(1) if match else None


def _tag(element) -> str:
    return _tag_name(element.tag)


def _tag_name(name: str) -> str:
    """``{namespace}sheet`` → ``sheet``."""
    return name.rsplit("}", 1)[-1]


def _xml(zf: zipfile.ZipFile, part: str):
    try:
        return ElementTree.fromstring(zf.read(part))
    except (KeyError, ElementTree.ParseError):
        return None


def _rels(zf: zipfile.ZipFile, part: str, kind: str | None = None) -> dict[str, str]:
    """``{relationship id: part path}`` for one part, optionally by type."""
    base, _, name = part.rpartition("/")
    root = _xml(zf, f"{base}/_rels/{name}.rels")
    if root is None:
        return {}
    out: dict[str, str] = {}
    for rel in root:
        rid, target = rel.get("Id"), rel.get("Target")
        if not rid or not target:
            continue
        if kind is not None and not rel.get("Type", "").endswith("/" + kind):
            continue
        out[rid] = _resolve(base, target)
    return out


def _resolve(base: str, target: str) -> str:
    """Resolve a relationship target against the part that declares it."""
    if target.startswith("/"):
        return target.lstrip("/")
    segments = base.split("/") if base else []
    for segment in target.split("/"):
        if segment == "..":
            if segments:
                segments.pop()
        elif segment not in ("", "."):
            segments.append(segment)
    return "/".join(segments)
