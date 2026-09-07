"""Boot the formualizer engine, and make it survive a broken reference.

Extracted mechanically from analyzer.py: instantiate the engine from the
workbook bytes, run the whole-workbook evaluation, and — when that fails
because one formula names something the engine cannot resolve — quarantine
the offending cells and retry, so the rest of the file keeps its computed
values.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from xml.sax.saxutils import unescape

import formualizer as fz

from linexcel.decompose import SCRATCH_SHEET
from linexcel.refs import col_to_num

# Guards to stay responsive on large workbooks.
SCAN_CHUNK_ROWS = 20_000
#: Ceiling on one ``get_formulas`` call, in cells. The engine hands back a
#: dense grid of Python strings, so a 16,384-column sheet read 20,000 rows at
#: a time would materialize 327 million of them in one go.
SCAN_CHUNK_CELLS = 1_000_000

#: A bracketed group: an external workbook (``[Budget.xlsx]Sheet1!A1``) or a
#: structured table reference (``SalesTable[Revenue]``). Neither resolves to a
#: cell the engine holds.
_BRACKETED_RE = re.compile(r"\[[^\]]*\]")
#: A 3-D sheet span — ``'First:Last'!A1`` or ``First:Last!A1``. The engine reads
#: the span as one sheet name and reports it missing.
_SPAN_RE = re.compile(r"'[^']*:[^']*'!|(?<![\w$)])[A-Za-z_][\w.]*:[A-Za-z_][\w.]*!")
#: Error guards. A formula using one may well have a defined value despite a
#: reference that cannot be resolved, so it is never isolated.
_GUARD_RE = re.compile(r"\b(?:IFERROR|IFNA|ISERROR|ISERR|ISNA)\s*\(", re.IGNORECASE)
#: A sheet qualifier, quoted or bare. The lookbehind keeps ``#REF!`` out of the
#: bare form: that is an error literal, not a sheet called REF.
_SHEET_QUALIFIER_RE = re.compile(r"'((?:[^']|'')+)'!|(?<![#\w.$])([A-Za-z_][\w.]*)!")

#: A cell element and its body. ``<c>`` never nests another ``<c>``, so the
#: lazy ``.*?</c>`` is safe; self-closing ``<c/>`` elements carry no formula.
_CELL_RE = re.compile(rb'<c\b[^>]*\br="([A-Z]{1,3})(\d+)"[^>]*>(.*?)</c>', re.S)
#: The ``<f>`` inside a cell body: either ``<f …/>`` (a shared-formula slave)
#: or ``<f …>text</f>``.
_F_RE = re.compile(rb"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
#: The shared-formula index, ``si="3"``.
_SI_RE = re.compile(rb'\bsi="(\d+)"')
_XML_ENTITIES = {"&quot;": '"', "&apos;": "'"}


@dataclass
class EngineSession:
    engine: object
    engine_sheets: set[str]
    engine_alive: bool
    quarantined: dict[tuple[str, int, int], str]
    scratch_ready: bool


def boot_engine(
    data: bytes,
    warnings: list[str],
) -> EngineSession:
    """Instantiate the engine and run its whole-workbook evaluation.

    ``evaluate_all`` is all-or-nothing, and it gives up on the *first*
    reference it cannot resolve — so a single formula pointing at another
    workbook costs every other cell in the file its computed value. When that
    happens, the offending formulas are cut out of the package bytes and the
    pass retried on the sanitized copy, leaving only them to the slower
    per-cell recovery.
    """
    engine = fz.Workbook.from_bytes(data)
    engine_sheets = set(engine.sheet_names)
    engine_alive = True
    quarantined: dict[tuple[str, int, int], str] = {}
    try:
        engine.evaluate_all()
    except Exception as exc:  # graph remains useful without values
        quarantined = _find_unresolvable(data, engine_sheets)
        retried = False
        if quarantined:
            try:
                engine = fz.Workbook.from_bytes(
                    _blank_formulas_in_package(data, quarantined)
                )
                engine.evaluate_all()
                retried = True
            except Exception:
                pass
        if not retried:
            # A failed global evaluation does not just drop the values: the
            # engine then reports no formula at all, which would leave the
            # graph empty. Rebuilding from the bytes gives the formulas back.
            engine = fz.Workbook.from_bytes(data)
        if retried:
            warnings.append(
                f"Global evaluation completed after isolating {len(quarantined)} "
                f"cell(s) whose references the engine cannot resolve; every other "
                f"cell was recomputed. Only those keep the value stored in the "
                f"file, if any. First blocker: {exc}"
            )
        else:
            warnings.append(f"Global evaluation incomplete: {exc}")
            # Values are recovered cell by cell further down.
            engine_alive = False
            quarantined = {}

    scratch_ready = _ensure_scratch(engine)
    return EngineSession(
        engine, engine_sheets, engine_alive, quarantined, scratch_ready
    )


def _find_unresolvable(
    data: bytes, engine_sheets: set[str]
) -> dict[tuple[str, int, int], str]:
    """The formulas that stop ``evaluate_all``, keyed by their cell.

    Only ever called after a global evaluation has already failed, so both ways
    of being wrong are safe: quarantining a formula that would in fact have
    evaluated costs it the same per-cell recovery every cell was getting anyway,
    and missing one simply leaves the retry failing as before.

    The suspects are read from the sheet XML, never from the engine. Reading
    them with ``get_formulas`` would force the engine to build its dependency
    graph one cell at a time — measured quadratic, about 18 s for 20,000 cells,
    i.e. half an hour for the kind of workbook that lands here — and that graph
    build is exactly what the failed ``evaluate_all`` proved impossible. The
    XML scan streams the parts the file already carries and stays linear.
    """
    suspects: dict[tuple[str, int, int], str] = {}
    for sheet, row, col, formula in _iter_formulas_xml(data):
        if sheet in engine_sheets and _is_unresolvable(formula, engine_sheets):
            suspects[(sheet, row, col)] = formula
    return suspects


def _iter_formulas_xml(data: bytes):
    """Yield ``(sheet, row, col, formula)`` for every formula cell in the package.

    The text is the file's own, ``=``-prefixed to match what the engine hands
    back from ``get_formulas``, with XML entities resolved. Shared formulas
    (``<f t="shared" si="n"/>``) yield their master's text untranslated: the
    relative references a translation would shift are exactly the ones no
    suspect check looks at, and the sheet qualifiers and brackets it does look
    at are identical in every cell of the share group.
    """
    from linexcel.loader import _parse_sheet_targets

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
            rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
            for name, zpath in _parse_sheet_targets(wb_xml, rels).items():
                if not zpath.endswith(".xml"):
                    continue
                try:
                    sheet_xml = zf.read(zpath)
                except KeyError:
                    continue
                shared: dict[bytes, bytes] = {}
                for cell in _CELL_RE.finditer(sheet_xml):
                    f = _F_RE.search(cell.group(3))
                    if f is None:
                        continue
                    text = f.group(2)
                    if text is None:
                        # A shared-formula slave: same formula as its master.
                        si = _SI_RE.search(f.group(1))
                        text = shared.get(si.group(1)) if si else None
                        if text is None:
                            continue
                    elif b't="shared"' in f.group(1):
                        si = _SI_RE.search(f.group(1))
                        if si:
                            shared[si.group(1)] = text
                    formula = "=" + unescape(
                        text.decode("utf-8", "replace"), _XML_ENTITIES
                    )
                    yield (
                        name,
                        int(cell.group(2)),
                        col_to_num(cell.group(1).decode("ascii")),
                        formula,
                    )
    except Exception:
        # A package this best-effort scan cannot read leaves the retry to fail
        # as it did before quarantine existed.
        return


def _blank_formulas_in_package(
    data: bytes, quarantined: dict[tuple[str, int, int], str]
) -> bytes:
    """Copy the package with the quarantined cells' ``<f>`` elements cut out.

    Blanking through the engine (``set_value`` / ``set_values_batch``) costs it
    a graph notice per cell — measured near 1.5 ms each even batched, so the
    tens of thousands of external-reference cells a real workbook lands here
    with meant minutes spent mutating. Cutting the elements out of the XML is a
    linear pass, and a cell that keeps its cached ``<v>`` comes back as the
    constant the user last saw, which is exactly the fallback value the
    quarantine used to throw away.

    A shared-formula master carries the text its slaves only name by index, so
    a group is cut whole or not at all; the suspect scan reads every slave as
    its master's text, which guarantees that.
    """
    from linexcel.loader import _parse_sheet_targets

    doomed: dict[str, set[tuple[int, int]]] = {}
    for sheet, row, col in quarantined:
        doomed.setdefault(sheet, set()).add((row, col))

    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
        path_of = _parse_sheet_targets(wb_xml, rels)
        sheet_of_path = {path: name for name, path in path_of.items()}
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for entry in zf.infolist():
                payload = zf.read(entry.filename)
                cells = doomed.get(sheet_of_path.get(entry.filename, ""))
                if cells:

                    def strip(m: re.Match[bytes], cells=cells) -> bytes:
                        key = (int(m.group(2)), col_to_num(m.group(1).decode("ascii")))
                        if key not in cells:
                            return m.group(0)
                        return _F_RE.sub(b"", m.group(0), count=1)

                    payload = _CELL_RE.sub(strip, payload)
                zout.writestr(entry, payload)
    return out.getvalue()


def _is_unresolvable(formula: str, engine_sheets: set[str]) -> bool:
    """Whether ``formula`` names something the engine cannot resolve.

    A guarded formula is never isolated, however broken its reference looks.
    ``IFERROR(NOSHEET!A1, 456)`` *has* a correct value — 456 — and blanking it
    does not merely cost that cell its own value: every range that spans it
    silently loses a term, so a `SUM` over the column returns a different
    number. Quarantine is only safe where there was no value to lose.
    """
    if _GUARD_RE.search(formula):
        return False
    if _BRACKETED_RE.search(formula) or _SPAN_RE.search(formula):
        return True
    # A sheet qualifier naming a sheet the engine never loaded.
    for match in _SHEET_QUALIFIER_RE.finditer(formula):
        name = match.group(1)
        name = name.replace("''", "'") if name else match.group(2)
        if name and name not in engine_sheets:
            return True
    return False


def _ensure_scratch(engine) -> bool:
    try:
        engine.add_sheet(SCRATCH_SHEET)
        return True
    except Exception:
        return SCRATCH_SHEET in set(engine.sheet_names)


def _chunk_rows(max_col: int) -> int:
    """How many rows to read per ``get_formulas`` call on a sheet that wide.

    Rows alone are the wrong unit: the engine returns a dense grid of Python
    strings, so 20,000 rows of a 16,384-column sheet is 327 million of them in
    a single call. Never zero — one row at a time is the floor, however wide.
    """
    return max(1, min(SCAN_CHUNK_ROWS, SCAN_CHUNK_CELLS // max(max_col, 1)))
