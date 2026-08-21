"""Workbooks a file depends on, and how far linexcel can follow them.

A formula reading ``'[1]Annual'!B4`` names another workbook. The file itself
says which one: ``xl/externalLinks/externalLink1.xml`` and its relationship
carry the path Excel last saw it at, and — when Excel saved the link — a cache
of the values that were read. Neither is visible to the calculation engine, so
without this module such a cell is a grey "external reference" node and every
formula above it loses its value.

Three levels of answer, in the order they are tried:

1. **Named.** The path and file name are always readable, so the report can say
   *which* workbook a cell depends on instead of showing ``[1]``.
2. **Cached.** Most files carry the values Excel last read across the link.
   They are what the user saw on screen, and they cost nothing to read.
3. **Resolved.** Given a folder of the referenced files, the workbook is opened
   and read for real — which is also where VBA lives when the code sits in an
   add-in (``.xlam``, ``.xla``) rather than in the workbook itself.

Level 3 is opt-in: ``analyze(..., refs_dir=...)``. Nothing is read from disk
unless a caller names the folder.
"""

from __future__ import annotations

import datetime
import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from xml.etree import ElementTree

#: Suffixes searched in a reference folder. The macro-enabled ones are also
#: where an add-in keeps the VBA a workbook calls into.
WORKBOOK_SUFFIXES = (".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlam", ".xla")
#: Suffixes that can carry a VBA project.
MACRO_SUFFIXES = (".xlsm", ".xlsb", ".xlam", ".xla", ".xltm")
#: How deep a reference folder is walked. One level of subfolders covers the
#: usual "everything next to the workbook, add-ins in a subfolder" layout
#: without turning a pointed argument into a disk crawl.
MAX_FOLDER_DEPTH = 2
#: Ceiling on how many cells are read out of one referenced workbook.
MAX_EXTERNAL_CELLS = 200_000

_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
#: ``'[1]Annual'!B4``, ``[1]Annual!B4`` and the full-path form Excel writes when
#: the link is not registered, ``'C:\\dir\\[Book.xlsx]Sheet'!A1``. The book is
#: either an index into the external links or a file name.
_EXTERNAL_REF_RE = re.compile(
    r"""(?P<quote>')?                 # the quote opens before the bracket
        (?P<dir>(?:[A-Za-z]:)?[^'\[\]!]*[\\/])?   # a path, ending on its separator
        \[(?P<book>[^\]]+)\]          # [1] or [Budget FY26.xlsx]
        (?P<sheet>(?:[^'!]|'')*)      # sheet name, up to the closing quote
        (?(quote)')                   # closed only if it was opened
        !(?P<cell>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)""",
    re.VERBOSE,
)


@dataclass
class ExternalBook:
    """One workbook the analyzed file reads from."""

    #: What formulas call it: ``1`` for ``[1]``, or the file name.
    key: str
    #: The path as the file declares it, unescaped.
    target: str
    #: File name alone, which is what a reference folder is searched for.
    name: str
    #: Sheet names, in the order the link declares them.
    sheets: list[str] = field(default_factory=list)
    #: Values Excel cached across the link, by ``(sheet, row, col)``.
    cached: dict[tuple[str, int, int], Any] = field(default_factory=dict)
    #: The file itself, once found in a reference folder.
    path: Path | None = None
    #: Values read from that file, by ``(sheet, row, col)``.
    values: dict[tuple[str, int, int], Any] = field(default_factory=dict)

    @property
    def resolved(self) -> bool:
        """True when the referenced workbook was actually read."""
        return self.path is not None

    def value(self, sheet: str, row: int, col: int) -> tuple[Any, str | None]:
        """``(value, source)`` for one cell of this workbook.

        The file on disk wins over the cache: the cache is what Excel read the
        last time it opened the link, which may be older than the workbook a
        caller has just pointed linexcel at.
        """
        key = (sheet, row, col)
        if key in self.values:
            return self.values[key], "external"
        if key in self.cached:
            return self.cached[key], "external-cache"
        return None, None


@dataclass
class ExternalRef:
    """An external reference as written in a formula."""

    #: The whole matched text, so it can be substituted back out.
    text: str
    book: str
    sheet: str
    cell: str
    #: The directory the formula spells out, when it does.
    directory: str = ""


def parse_external_refs(formula: str) -> list[ExternalRef]:
    """Every external reference a formula makes, in order of appearance."""
    refs = []
    for match in _EXTERNAL_REF_RE.finditer(formula):
        refs.append(
            ExternalRef(
                text=match.group(0),
                book=match.group("book").strip(),
                sheet=match.group("sheet").replace("''", "'"),
                cell=match.group("cell").replace("$", ""),
                directory=match.group("dir") or "",
            )
        )
    return refs


def read_external_links(data: bytes) -> dict[str, ExternalBook]:
    """The workbooks a file declares a link to, keyed by ``[N]`` index.

    The index is the position in ``<externalReferences>`` in ``xl/workbook.xml``
    — not the number in the part name, which only coincides most of the time.
    """
    books: dict[str, ExternalBook] = {}
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            book = _xml(zf, "xl/workbook.xml")
            if book is None:
                return books
            rels = _rels(zf, "xl/workbook.xml")
            index = 0
            for element in book.iter():
                if _tag(element) != "externalReference":
                    continue
                index += 1
                rid = next(
                    (v for k, v in element.attrib.items() if _tag_of(k) == "id"), None
                )
                part = rels.get(rid or "", (None, None))[1]
                if part is None:
                    continue
                entry = _read_link_part(zf, part, str(index))
                if entry is not None:
                    books[str(index)] = entry
    except Exception:
        return {}
    return books


def find_workbooks(folder: Path) -> dict[str, Path]:
    """Workbook files in ``folder``, keyed by lowercased file name.

    The first match wins, and the shallowest one is seen first, so a file next
    to the workbook takes precedence over a copy buried in a subfolder.
    """
    found: dict[str, Path] = {}
    if not folder.is_dir():
        return found
    for depth in range(MAX_FOLDER_DEPTH):
        pattern = "/".join(["*"] * (depth + 1))
        for path in sorted(folder.glob(pattern)):
            if not path.is_file() or path.suffix.lower() not in WORKBOOK_SUFFIXES:
                continue
            found.setdefault(path.name.lower(), path)
    return found


def resolve_books(
    books: dict[str, ExternalBook], folder: Path, warnings: list[str]
) -> None:
    """Read every declared workbook that the folder actually holds."""
    available = find_workbooks(folder)
    for entry in books.values():
        path = available.get(entry.name.lower())
        if path is None:
            warnings.append(
                f"External workbook '{entry.name}' is not in the reference "
                f"folder; its cells keep the value cached in the file, if any"
            )
            continue
        try:
            entry.values = read_workbook_values(path)
            entry.path = path
        except Exception as exc:
            warnings.append(
                f"External workbook '{entry.name}' could not be read: {exc}"
            )


def read_workbook_values(path: Path) -> dict[tuple[str, int, int], Any]:
    """Every value of a workbook, by ``(sheet, row, col)``.

    Only values: a referenced workbook is read for what the formulas above it
    need, and Excel itself stores nothing but values across a link.
    """
    from python_calamine import CalamineWorkbook

    # Imported here rather than at module scope: analyzer imports this module,
    # so the dependency only goes the other way once the call is under way.
    from linexcel.analyzer import MAX_DENSE_CELLS, declared_cells

    # calamine builds a sheet as a dense rows × columns array before handing
    # anything back, so a workbook declaring A1:XFD1048576 asks the allocator
    # for 512 GiB and *aborts the process* — a Rust allocation failure, not an
    # exception, so the caller's try/except would never see it. A file that
    # claims more than any sheet can hold is refused by name instead.
    declared = declared_cells(path.read_bytes())
    if declared > MAX_DENSE_CELLS:
        raise ValueError(
            f"it declares a used range of {declared:,} cells, more than can be "
            f"read; open it, delete the empty rows below and columns right of "
            f"the data, and save"
        )

    values: dict[tuple[str, int, int], Any] = {}
    workbook = CalamineWorkbook.from_path(str(path))
    for name in workbook.sheet_names:
        rows = workbook.get_sheet_by_name(name).to_python(skip_empty_area=False)
        scanned = 0
        for r_index, row in enumerate(rows):
            scanned += len(row)
            if scanned > MAX_EXTERNAL_CELLS:
                break
            for c_index, value in enumerate(row):
                if value is None or value == "":
                    continue
                if isinstance(value, datetime.date) and not isinstance(
                    value, datetime.datetime
                ):
                    value = datetime.datetime(value.year, value.month, value.day)
                values[(name, r_index + 1, c_index + 1)] = value
    return values


def macro_files(folder: Path) -> list[Path]:
    """Files in ``folder`` that can carry a VBA project."""
    return [
        path
        for name, path in sorted(find_workbooks(folder).items())
        if path.suffix.lower() in MACRO_SUFFIXES
    ]


# -- package plumbing -------------------------------------------------------


def _read_link_part(zf: zipfile.ZipFile, part: str, key: str) -> ExternalBook | None:
    root = _xml(zf, part)
    if root is None:
        return None
    target = ""
    for rid, (rel_type, resolved) in _rels(zf, part).items():
        if rel_type.endswith("/externalLinkPath"):
            target = resolved
            break
    sheets = [
        element.get("val") or ""
        for element in root.iter()
        if _tag(element) == "sheetName"
    ]
    entry = ExternalBook(
        key=key,
        target=target,
        name=target.replace("\\", "/").rsplit("/", 1)[-1] or target,
        sheets=sheets,
    )
    _read_cached_cells(root, sheets, entry.cached)
    return entry


def _read_cached_cells(
    root, sheets: list[str], into: dict[tuple[str, int, int], Any]
) -> None:
    """The values Excel last read across the link, if it saved them.

    ``sheetId`` indexes the ``<sheetNames>`` list declared alongside, and each
    cell carries its type the same way a worksheet does — ``str`` for text,
    ``b`` for a boolean, ``e`` for an error, nothing for a number.
    """
    for sheet_data in root.iter():
        if _tag(sheet_data) != "sheetData":
            continue
        try:
            sheet_index = int(sheet_data.get("sheetId") or 0)
        except ValueError:
            continue
        sheet = sheets[sheet_index] if sheet_index < len(sheets) else ""
        for cell in sheet_data.iter():
            if _tag(cell) != "cell":
                continue
            address = cell.get("r")
            if not address:
                continue
            position = _a1_to_rowcol(address)
            if position is None:
                continue
            text = next(
                (child.text for child in cell if _tag(child) == "v"),
                None,
            )
            if text is None:
                continue
            into[(sheet, *position)] = _typed(text, cell.get("t"))


def _typed(text: str, kind: str | None) -> Any:
    if kind in ("str", "e"):
        return text
    if kind == "b":
        return text not in ("0", "false", "FALSE")
    try:
        return float(text)
    except ValueError:
        return text


_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def _a1_to_rowcol(address: str) -> tuple[int, int] | None:
    match = _A1_RE.match(address.strip())
    if match is None:
        return None
    column = 0
    for char in match.group(1).upper():
        column = column * 26 + (ord(char) - 64)
    return int(match.group(2)), column


def _tag(element) -> str:
    return _tag_of(element.tag)


def _tag_of(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _xml(zf: zipfile.ZipFile, part: str):
    try:
        return ElementTree.fromstring(zf.read(part))
    except (KeyError, ElementTree.ParseError):
        return None


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """``{id: (type, target)}`` for one part, external targets kept verbatim."""
    base, _, name = part.rpartition("/")
    root = _xml(zf, f"{base}/_rels/{name}.rels")
    if root is None:
        return {}
    out: dict[str, tuple[str, str]] = {}
    for rel in root:
        rid, target = rel.get("Id"), rel.get("Target")
        if not rid or not target:
            continue
        target = unquote(target)
        if rel.get("TargetMode") != "External":
            target = _resolve(base, target)
        out[rid] = (rel.get("Type", ""), target)
    return out


def _resolve(base: str, target: str) -> str:
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
