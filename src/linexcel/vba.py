"""Extraction and static lineage of embedded VBA code.

Extraction relies on oletools (olevba). Analysis is static and deliberately
heuristic: it identifies procedures, the internal call graph, and cell/range
access expressed literally (Range("A1"), Cells(2, 3), [A1:B4],
Worksheets("X").Range(...)).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_PROC_RE = re.compile(
    r"^[ \t]*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
    re.IGNORECASE | re.MULTILINE,
)
_END_PROC_RE = re.compile(
    r"^[ \t]*End\s+(Sub|Function|Property)\b", re.IGNORECASE | re.MULTILINE
)

# Worksheets("Name").  /  Sheets("Name").  — captures sheet name
_SHEET_QUAL = (
    r"(?:(?:ThisWorkbook|ActiveWorkbook|\w+)\.)?"
    r"(?:Worksheets|Sheets)\(\s*\"(?P<sheet>[^\"]+)\"\s*\)\s*\."
)
# Range("A1") / Range("A1:B2") / Range("A1", "B2")
_RANGE_CALL = r"Range\(\s*\"(?P<a1>[^\"]+)\"\s*(?:,\s*\"(?P<a2>[^\"]+)\"\s*)?\)"
# Cells(2, 3) with literal arguments only
_CELLS_CALL = (
    r"Cells\(\s*(?P<row>\d+)\s*,"
    r"\s*(?P<colq>\"?)(?P<col>[A-Za-z]{1,3}|\d+)(?P=colq)\s*\)"
)

_REF_RE = re.compile(
    rf"(?:{_SHEET_QUAL})?(?:(?P<range>{_RANGE_CALL})|(?P<cells>{_CELLS_CALL}))",
    re.IGNORECASE,
)
# Shortcut [A1] / [A1:B2]
_BRACKET_RE = re.compile(
    r"\[(?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)\]"
)

_COMMENT_RE = re.compile(r"(?<!\")'.*$", re.MULTILINE)
_STRINGS_KEEP = re.compile(r'"[^"]*"')


@dataclass
class VbaRef:
    """A range access detected in a procedure."""

    sheet: str | None
    ref: str
    access: str  # "read" | "write"
    line: int


@dataclass
class VbaProc:
    module: str
    name: str
    kind: str
    line_start: int
    line_end: int
    code: str
    refs: list[VbaRef] = field(default_factory=list)
    calls: list[str] = field(default_factory=list)


def extract_vba_modules(data: bytes, filename: str) -> dict[str, str]:
    """Extract {module_name: code} via olevba. Empty dict if no VBA."""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:  # pragma: no cover - dependency installed in prod
        return {}
    try:
        parser = VBA_Parser(filename, data=data)
    except Exception:
        return {}
    modules: dict[str, str] = {}
    try:
        if parser.detect_vba_macros():
            for _f, _path, vba_filename, code in parser.extract_macros():
                name = (vba_filename or "Module").rsplit("/", 1)[-1]
                name = re.sub(r"\.(bas|cls|frm)$", "", name, flags=re.IGNORECASE)
                if code and code.strip():
                    existing = modules.get(name)
                    modules[name] = (existing + "\n" + code) if existing else code
    except Exception:
        return {}
    finally:
        try:
            parser.close()
        except Exception:
            pass
    return modules


def _strip_comments(line: str) -> str:
    """Strip a trailing comment while respecting string literals."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    return line


def _split_procedures(module: str, code: str) -> list[VbaProc]:
    lines = code.splitlines()
    procs: list[VbaProc] = []
    current: VbaProc | None = None
    body: list[str] = []
    for idx, raw in enumerate(lines, start=1):
        line = _strip_comments(raw)
        m = _PROC_RE.match(line)
        if m and current is None:
            kind = re.sub(r"\s+", " ", m.group(1)).title()
            current = VbaProc(module, m.group(2), kind, idx, idx, "")
            body = [raw]
            continue
        if current is not None:
            body.append(raw)
            if _END_PROC_RE.match(line):
                current.line_end = idx
                current.code = "\n".join(body)
                procs.append(current)
                current = None
                body = []
    if current is not None:
        current.line_end = len(lines)
        current.code = "\n".join(body)
        procs.append(current)
    return procs


def _detect_access(line: str, match_end: int) -> str:
    """Write if the reference is followed by an assignment at statement level."""
    rest = line[match_end:]
    # skip members (.Value, .Formula, .Offset(...)) after the reference
    rest = re.sub(r"^(?:\.\w+(?:\([^()]*\))?)*", "", rest).lstrip()
    if rest.startswith("=") and not rest.startswith("=="):
        return "write"
    return "read"


def _find_refs(proc: VbaProc) -> list[VbaRef]:
    refs: list[VbaRef] = []
    for offset, raw in enumerate(proc.code.splitlines()):
        line = _strip_comments(raw)
        for m in _REF_RE.finditer(line):
            sheet = m.group("sheet")
            if m.group("range"):
                a1, a2 = m.group("a1"), m.group("a2")
                ref = f"{a1}:{a2}" if a2 else a1
            else:
                col = m.group("col")
                col_txt = col if col.isalpha() else _num_col(int(col))
                ref = f"{col_txt}{m.group('row')}"
            access = _detect_access(line, m.end())
            refs.append(VbaRef(sheet, ref, access, proc.line_start + offset))
        for m in _BRACKET_RE.finditer(line):
            access = _detect_access(line, m.end())
            refs.append(VbaRef(None, m.group("ref"), access, proc.line_start + offset))
    return refs


def _num_col(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _find_calls(proc: VbaProc, known: set[str]) -> list[str]:
    calls: set[str] = set()
    body = "\n".join(_strip_comments(line) for line in proc.code.splitlines()[1:])
    body = _STRINGS_KEEP.sub('""', body)
    for m in re.finditer(r"\b([A-Za-z_]\w*)\b", body):
        name = m.group(1)
        if name != proc.name and name in known:
            calls.add(name)
    return sorted(calls)


def analyze_vba(modules: dict[str, str]) -> list[VbaProc]:
    """Full static analysis: procedures, range access, call graph."""
    procs: list[VbaProc] = []
    for module, code in modules.items():
        procs.extend(_split_procedures(module, code))
    known = {p.name for p in procs}
    for proc in procs:
        proc.refs = _find_refs(proc)
        proc.calls = _find_calls(proc, known)
    return procs
