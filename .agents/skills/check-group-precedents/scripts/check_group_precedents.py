#!/usr/bin/env python3
"""Exhaustive bounded check of the precedents of a formula group.

Property (a group's precedents are what its cells read). linexcel folds the
cells of a sheet whose formulas share one R1C1 key into a single group node,
and wires that node to the rectangle obtained by stretching a reference of one
cell, the representative, over the group's bounding box. For every group:

- complete: the rectangle each cell's own reference reads lies inside the
  rectangle the node claims. Otherwise a precedent is silently missing;
- exact: the claimed rectangle is the smallest one holding what the cells
  read. Otherwise the node shows a precedent that no cell reads.

Supporting property (keys separate different references). At one cell, two
different references never produce the same key. Without it a group could mix
formulas that do not carry the same logic.

The check runs the real code. Nothing is modelled by hand except the oracle:
each reference is generated from integers, so the rectangle it reads is known
by construction, without parsing.

    uv run python <this file>                 the check
    uv run python <this file> --mutants       prove the check can fail: it breaks
                                              refs.py in memory, six ways
    uv run python <this file> --show CLASS    every counterexample of one class
    uv run python <this file> --update-known  rewrite known_findings.json

Exit codes: 0 no violation; 3 only the findings recorded in
known_findings.json; 1 a violation that is not recorded there, a stale model
of the call site, or (with --mutants) a broken variant that went unnoticed.
"""

from __future__ import annotations

import argparse
import itertools
import json
import sys
import time
import types
from contextlib import contextmanager
from dataclasses import dataclass
from importlib.metadata import version
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
KNOWN = Path(__file__).with_name("known_findings.json")
SHEET = "S"

# The composition under test, as written at its call sites. If one of these
# lines disappears, the check no longer describes the code: stop and re-read.
CALL_SITES = {
    "src/linexcel/graph.py": [
        "detail = parse_ref_detailed(ref, default_sheet=sheet)",
        "stretch_ref(detail, rep_r, rep_c, (rmin, rmax), (cmin, cmax))",
        "rmin, cmin, rmax, cmax = grp.bbox",
        "rep_r, rep_c = grp.rep",
    ],
    "src/linexcel/sweep.py": [
        "key = (sheet, canonical_r1c1(formula, r, c))",
        "return min(self.cells)",
        "return min(rows), min(cols), max(rows), max(cols)",
    ],
    "src/linexcel/rewrite.py": ["conv = ref_to_r1c1(v, row, col)"],
}

# Broken variants of src/linexcel/refs.py. Each replaces one piece of text and
# must make the check report a violation that the unbroken code does not.
MUTANTS = [
    (
        "stretch_ref moves the end row by the smallest displacement",
        "rect.r2 + (rows_span[1] - rep_row)",
        "rect.r2 + (rows_span[0] - rep_row)",
    ),
    (
        "stretch_ref ignores the column span for the start column",
        "rect.c1 + (cols_span[0] - rep_col)",
        "rect.c1",
    ),
    (
        "stretch_ref treats an anchored end row as relative",
        "r2 = rect.r2 if detail.row2_abs else",
        "r2 = rect.r2 if False else",
    ),
    (
        "the key writes an anchored row like a relative one",
        'r = f"R{row}" if abs_row else _rel("R", row - base_row)',
        'r = _rel("R", row - base_row)',
    ),
    (
        "the key drops the column of a cell reference",
        "        return r + c\n",
        "        return r\n",
    ),
    (
        "parse_ref_detailed reads the row anchor of a cell as its column anchor",
        'return row, col, m.group(3) == "$", m.group(1) == "$"',
        'return row, col, m.group(1) == "$", m.group(1) == "$"',
    ),
]


def letters(col: int) -> str:
    """Column letters, written here so the oracle does not use linexcel."""
    out = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


@dataclass(frozen=True)
class Ref:
    """One generated reference and the rectangle it reads, known by construction."""

    text: str
    rect: tuple[int, int, int, int]  # r1, c1, r2, c2, inclusive
    shape: str  # cell | range | columns | rows
    anchors: str  # template such as "A1:A$1": which bounds carry a $
    reversed_rows: bool
    reversed_cols: bool


def generate_refs(rows: list[int], cols: list[int], max_row: int, max_col: int):
    flags = (False, True)
    cells = [(r, c) for r in rows for c in cols]

    def endpoint(r, c, rabs, cabs):
        return f"{'$' * cabs}{letters(c)}{'$' * rabs}{r}"

    def template(c1abs, r1abs, c2abs=None, r2abs=None, kind="cell"):
        a = f"{'$' * c1abs}A{'$' * r1abs}1"
        if kind == "cell":
            return a
        return f"{a}:{'$' * c2abs}A{'$' * r2abs}1"

    for r, c in cells:
        for rabs, cabs in itertools.product(flags, flags):
            yield Ref(
                endpoint(r, c, rabs, cabs),
                (r, c, r, c),
                "cell",
                template(cabs, rabs),
                False,
                False,
            )
    for (r1, c1), (r2, c2) in itertools.product(cells, cells):
        rect = (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
        for r1a, c1a, r2a, c2a in itertools.product(flags, repeat=4):
            yield Ref(
                f"{endpoint(r1, c1, r1a, c1a)}:{endpoint(r2, c2, r2a, c2a)}",
                rect,
                "range",
                template(c1a, r1a, c2a, r2a, "range"),
                r1 > r2,
                c1 > c2,
            )
    for c1, c2 in itertools.product(cols, cols):
        for c1a, c2a in itertools.product(flags, flags):
            yield Ref(
                f"{'$' * c1a}{letters(c1)}:{'$' * c2a}{letters(c2)}",
                (1, min(c1, c2), max_row, max(c1, c2)),
                "columns",
                f"{'$' * c1a}A:{'$' * c2a}A",
                False,
                c1 > c2,
            )
    for r1, r2 in itertools.product(rows, rows):
        for r1a, r2a in itertools.product(flags, flags):
            yield Ref(
                f"{'$' * r1a}{r1}:{'$' * r2a}{r2}",
                (min(r1, r2), 1, max(r1, r2), max_col),
                "rows",
                f"{'$' * r1a}1:{'$' * r2a}1",
                r1 > r2,
                False,
            )


@dataclass
class Violation:
    kind: str  # missing-precedent | extra-precedent | key-collision | not-resolved
    klass: str
    rank: tuple  # smaller is a simpler example to show first
    lines: list[str]


class Pipeline:
    """The real functions, in the order the analyzer calls them."""

    def __init__(self, refs_module=None):
        import formualizer as fz

        from linexcel import refs, rewrite
        from linexcel.decompose import typed_ast_dict
        from linexcel.graph import _collect_ref_strings
        from linexcel.sweep import FormulaGroup

        self.refs = refs_module or refs
        self.rewrite = rewrite
        self.fz = fz
        self.typed_ast_dict = typed_ast_dict
        self.collect = _collect_ref_strings
        self.group = FormulaGroup
        self._asts: dict[str, list[str]] = {}

    @contextmanager
    def active(self):
        """Make rewrite.py use this refs module, as its imports would."""
        names = ("parse_ref", "quote_sheet", "ref_to_r1c1", "split_sheet_prefix")
        saved = {name: getattr(self.rewrite, name) for name in names}
        for name in names:
            setattr(self.rewrite, name, getattr(self.refs, name))
        try:
            yield self
        finally:
            for name, value in saved.items():
                setattr(self.rewrite, name, value)

    def key(self, formula: str, row: int, col: int) -> str:
        # sweep.py: key = (sheet, canonical_r1c1(formula, r, c))
        return self.rewrite.canonical_r1c1(formula, row, col)

    def ast_refs(self, formula: str) -> list[str]:
        # graph.py: refs = _collect_ref_strings(typed_ast_dict(fz.parse(...)))
        found = self._asts.get(formula)
        if found is None:
            ast = self.fz.parse("=" + formula)
            found = self._asts[formula] = list(self.collect(self.typed_ast_dict(ast)))
        return found

    def claimed(self, formula: str, cells: list[tuple[int, int]]):
        """Rectangles the group node is wired to, or None when not resolved."""
        grp = self.group(SHEET, "", cells=list(cells))
        rep_r, rep_c = grp.rep
        rmin, cmin, rmax, cmax = grp.bbox
        out = []
        for ref in self.ast_refs(formula):
            detail = self.refs.parse_ref_detailed(ref, default_sheet=SHEET)
            if detail is None:
                return None
            rect = (
                self.refs.stretch_ref(detail, rep_r, rep_c, (rmin, rmax), (cmin, cmax))
                if len(cells) > 1
                else detail.rect
            )
            out.append((detail, rect))
        return out


def mechanism(detail, axis: str) -> str:
    """Why an axis can lose cells, read off the representative's reference.

    ``stretch_ref`` moves the first bound by the group's smallest displacement
    and the second by its largest. That is the union of what the cells read
    only when the first bound stays the low one in every cell.
    """
    if axis == "rows":
        first, second = detail.row1_abs, detail.row2_abs
        backwards = detail.rect.r1 > detail.rect.r2
    else:
        first, second = detail.col1_abs, detail.col2_abs
        backwards = detail.rect.c1 > detail.rect.c2
    if first != second:
        return "one end anchored, one end moving"
    if not first and backwards:
        return "both ends moving, written high to low"
    return "UNEXPLAINED " + ("anchored" if first else "moving") + " ends"


def addr(cell: tuple[int, int]) -> str:
    return f"{letters(cell[1])}{cell[0]}"


def box(rect: tuple[int, int, int, int]) -> str:
    r1, c1, r2, c2 = rect
    a, b = f"{letters(c1)}{r1}", f"{letters(c2)}{r2}"
    return a if a == b else f"{a}:{b}"


def check_window(pipe: Pipeline, rows, cols, prefix: str, max_group: int):
    """Every group of cells of the window, of every size, sharing a key."""
    refs_mod = pipe.refs
    refs = list(generate_refs(rows, cols, refs_mod.MAX_ROW, refs_mod.MAX_COL))
    cells = [(r, c) for r in rows for c in cols]
    wrap = "SUM({})".format
    violations: list[Violation] = []
    by_key: dict[str, dict[tuple[int, int], Ref]] = {}
    stats = {"formulas": 0, "keys": 0, "groups": 0, "memberChecks": 0}

    for cell in cells:
        seen: dict[str, Ref] = {}
        for ref in refs:
            stats["formulas"] += 1
            key = pipe.key(wrap(prefix + ref.text), *cell)
            other = seen.get(key)
            if other is not None:
                violations.append(
                    Violation(
                        "key-collision",
                        f"key-collision|{other.shape}/{ref.shape}",
                        (bool(prefix), False, False, False, False, 1, max(cell), cell),
                        [
                            f"cell {SHEET}!{addr(cell)}: two different references "
                            "get the same key",
                            f"  =SUM({prefix}{other.text})  and  "
                            f"=SUM({prefix}{ref.text})",
                            f"  key {key}",
                            "  cells holding either formula would be folded "
                            "into one group node",
                        ],
                    )
                )
                continue
            seen[key] = ref
            by_key.setdefault(key, {})[cell] = ref
            # A formula alone is not stretched: graph.py wires it to
            # `detail.rect` as parsed. That rectangle must be what it reads.
            alone = pipe.claimed(wrap(prefix + ref.text), [cell])
            rects = None if alone is None else [r for _d, r in alone]
            problem = judge(rects, ref, prefix_sheet(prefix))
            if problem is None and rects is not None:
                r = rects[0]
                if (r.r1, r.c1, r.r2, r.c2) != ref.rect:
                    problem = ("extra-precedent", ("rectangle",))
            if problem is not None:
                kind, axes = problem
                claims = (
                    "nothing (reference not resolved)"
                    if rects is None
                    else ", ".join(
                        f"{r.sheet}!rows {r.r1}-{r.r2}, columns "
                        f"{letters(r.c1)}-{letters(r.c2)}"
                        for r in rects
                    )
                )
                violations.append(
                    Violation(
                        kind,
                        f"{kind}|lone formula|{'+'.join(axes)}",
                        (
                            bool(prefix),
                            ref.shape != "range",
                            False,
                            ref.reversed_rows,
                            False,
                            1,
                            max(cell),
                            cell,
                        ),
                        [
                            f"cell {SHEET}!{addr(cell)} alone, not in a group",
                            f"  it holds =SUM({prefix}{ref.text}), "
                            f"which reads {box(ref.rect)}",
                            f"  its node is wired to: {claims}",
                            "  a rectangle whose first column is after its last "
                            "one contains no cell: the node gets no precedent"
                            if kind == "missing-precedent"
                            else "  the node is not wired to exactly what it reads",
                        ],
                    )
                )
    stats["keys"] = len(by_key)

    expected_sheet = prefix_sheet(prefix)
    for key, members in by_key.items():
        if len(members) < 2:
            continue
        ordered = sorted(members)
        shapes: set[tuple[tuple[int, int], tuple[int, int, int, int]]] = set()
        # A group of any size has the representative and bounding box of some
        # sub-group of at most four of its cells: the smallest cell, and cells
        # reaching the last row, the first column and the last column.
        for size in range(2, min(max_group, len(ordered)) + 1):
            for combo in itertools.combinations(ordered, size):
                rs = [r for r, _ in combo]
                cs = [c for _, c in combo]
                shapes.add((combo[0], (min(rs), min(cs), max(rs), max(cs))))
        for rep, bbox in shapes:
            stats["groups"] += 1
            rmin, cmin, rmax, cmax = bbox
            inside = [
                m
                for m in ordered
                if m >= rep and rmin <= m[0] <= rmax and cmin <= m[1] <= cmax
            ]
            # The same representative and box, rebuilt from real cells.
            witness = [rep, (rmax, cmin), (rmax, cmax), (rmin, cmax)]
            formula = wrap(prefix + members[rep].text)
            group_cells = sorted(set(inside))
            claimed = pipe.claimed(formula, group_cells)
            got = pipe.group(SHEET, "", cells=list(group_cells))
            if got.rep != rep or got.bbox != bbox:
                raise AssertionError(f"enumeration bug: {witness} {got.rep} {got.bbox}")
            # `inside` is the largest group with this representative and box.
            # Its claim must not reach past the smallest rectangle holding
            # what its cells read: an edge from a cell nobody reads is false
            # lineage too.
            if claimed is not None and len(claimed) == 1:
                detail, rect = claimed[0]
                read = [members[m].rect for m in inside]
                hull = (
                    min(r[0] for r in read),
                    min(r[1] for r in read),
                    max(r[2] for r in read),
                    max(r[3] for r in read),
                )
                extra = [
                    axis
                    for axis, low, high, lo, hi in (
                        ("rows", rect.r1, rect.r2, hull[0], hull[2]),
                        ("columns", rect.c1, rect.c2, hull[1], hull[3]),
                    )
                    if low < lo or high > hi
                ]
                if extra and rect.sheet == expected_sheet:
                    rep_ref = members[rep]
                    lines = [
                        f"group on sheet {SHEET}: cells "
                        + ", ".join(addr(c) for c in group_cells[:6])
                        + (" …" if len(group_cells) > 6 else "")
                        + f"  (representative {addr(rep)}, box {box(bbox)})",
                        f"  key shared by all of them: {key}",
                        f"  {addr(rep)} holds =SUM({prefix}{rep_ref.text})",
                        f"  together the cells read no further than {box(hull)}",
                        "  the group node is wired to: "
                        f"{rect.sheet}!{box((rect.r1, rect.c1, rect.r2, rect.c2))}",
                        "  so the node claims cells that none of its cells reads",
                    ]
                    rank = (
                        bool(prefix),
                        rep_ref.shape != "range",
                        False,
                        rep_ref.reversed_rows or rep_ref.reversed_cols,
                        False,
                        len(group_cells),
                        max(rmax, cmax),
                        rep,
                    )
                    for axis in extra:
                        violations.append(
                            Violation(
                                "extra-precedent",
                                f"extra-precedent|{axis}|{mechanism(detail, axis)}",
                                rank,
                                lines,
                            )
                        )
            for member in inside:
                stats["memberChecks"] += 1
                ref = members[member]
                rects = None if claimed is None else [r for _d, r in claimed]
                problem = judge(rects, ref, expected_sheet)
                if problem is None:
                    continue
                kind, axes = problem
                if kind == "missing-precedent" and axes != ("sheet",):
                    detail = claimed[0][0]
                    names = [f"{kind}|{a}|{mechanism(detail, a)}" for a in axes]
                else:
                    names = [f"{kind}|{'+'.join(axes)}"]
                lines = describe(
                    prefix, key, rep, bbox, group_cells, members, member, rects
                )
                rep_ref = members[rep]
                # Prefer an example a reader recognises: a plain range, written
                # low to high, that does not read the group's own cells.
                circular = any(
                    not (
                        m.rect[2] < rmin
                        or m.rect[0] > rmax
                        or m.rect[3] < cmin
                        or m.rect[1] > cmax
                    )
                    for m in (rep_ref, ref)
                )
                rank = (
                    bool(prefix),
                    rep_ref.shape != "range",
                    circular,
                    rep_ref.reversed_rows or rep_ref.reversed_cols,
                    member == rep,
                    len(group_cells),
                    max(rmax, cmax),
                    member,
                )
                for klass in names:
                    violations.append(Violation(kind, klass, rank, lines))
    return stats, violations


def prefix_sheet(prefix: str) -> str:
    if not prefix:
        return SHEET
    name = prefix[:-1]
    if name.startswith("'"):
        name = name[1:-1].replace("''", "'")
    return name


def judge(claimed, ref: Ref, sheet: str):
    """None when a claimed rectangle contains what the member reads."""
    if claimed is None:
        return "not-resolved", ("all",)
    r1, c1, r2, c2 = ref.rect
    rows_ok = cols_ok = False
    for rect in claimed:
        if rect.sheet != sheet:
            continue
        in_rows = rect.r1 <= r1 and r2 <= rect.r2
        in_cols = rect.c1 <= c1 and c2 <= rect.c2
        if in_rows and in_cols:
            return None
        rows_ok |= in_rows
        cols_ok |= in_cols
    if not any(rect.sheet == sheet for rect in claimed):
        return "missing-precedent", ("sheet",)
    axes = tuple(
        axis for axis, ok in (("rows", rows_ok), ("columns", cols_ok)) if not ok
    )
    return "missing-precedent", axes


def describe(prefix, key, rep, bbox, cells, members, member, claimed) -> list[str]:
    shown = ", ".join(addr(c) for c in cells[:6]) + (" …" if len(cells) > 6 else "")
    claims = (
        ", ".join(f"{r.sheet}!{box((r.r1, r.c1, r.r2, r.c2))}" for r in claimed)
        if claimed
        else "nothing (reference not resolved)"
    )
    reads = box(members[member].rect)
    lines = [
        f"group on sheet {SHEET}: cells {shown}  "
        f"(representative {addr(rep)}, box {box(bbox)})",
        f"  key shared by all of them: {key}",
    ]
    if member == rep:
        lines.append(
            f"  {addr(rep)} holds =SUM({prefix}{members[rep].text}), "
            f"which reads {reads}"
        )
    else:
        lines += [
            f"  {addr(rep)} holds =SUM({prefix}{members[rep].text})",
            f"  {addr(member)} holds =SUM({prefix}{members[member].text}), "
            f"which reads {reads}",
        ]
    return lines + [
        f"  the group node is wired to: {claims}",
        f"  so part of {reads} read by {addr(member)} has no edge to the group",
    ]


def windows(pipe: Pipeline, quick: bool):
    top, right = pipe.refs.MAX_ROW, pipe.refs.MAX_COL
    plan = [
        ("top-left corner, 4 rows x 3 columns", [1, 2, 3, 4], [1, 2, 3], ""),
        ("top-left corner, 3 rows x 4 columns", [1, 2, 3], [1, 2, 3, 4], ""),
    ]
    if not quick:
        plan += [
            (
                "bottom-right corner, 3 rows x 3 columns",
                [top - 2, top - 1, top],
                [right - 2, right - 1, right],
                "",
            ),
            (
                "another sheet, quoted name, 3 rows x 2 columns",
                [1, 2, 3],
                [1, 2],
                "'O''Brien 2'!",
            ),
        ]
    return plan


def run(pipe: Pipeline, quick: bool, max_group: int, verbose: bool):
    totals = {"formulas": 0, "keys": 0, "groups": 0, "memberChecks": 0}
    found: list[Violation] = []
    with pipe.active():
        for label, rows, cols, prefix in windows(pipe, quick):
            started = time.time()
            stats, violations = check_window(pipe, rows, cols, prefix, max_group)
            for name, value in stats.items():
                totals[name] += value
            found += violations
            if verbose:
                print(
                    f"  {label}: {stats['formulas']:,} formulas, "
                    f"{stats['keys']:,} keys, {stats['groups']:,} group shapes, "
                    f"{stats['memberChecks']:,} member checks, "
                    f"{len(violations):,} violations, "
                    f"{time.time() - started:.1f}s"
                )
    return totals, found


def classes(found: list[Violation]) -> dict[str, list[Violation]]:
    out: dict[str, list[Violation]] = {}
    for violation in found:
        out.setdefault(violation.klass, []).append(violation)
    for items in out.values():
        items.sort(key=lambda v: (v.rank, v.lines))
    return out


def stale_call_sites() -> list[str]:
    missing = []
    for name, needles in CALL_SITES.items():
        text = (ROOT / name).read_text(encoding="utf-8")
        missing += [f"{name}: {needle}" for needle in needles if needle not in text]
    return missing


class StaleMutantError(Exception):
    """The text a broken variant replaces is no longer in refs.py."""


def mutated_refs(old: str, new: str):
    import linexcel.refs  # noqa: F401 - registers the module replaced below

    path = ROOT / "src" / "linexcel" / "refs.py"
    source = path.read_text(encoding="utf-8")
    if source.count(old) != 1:
        raise StaleMutantError(
            f"expected exactly one occurrence in refs.py of: {old!r} "
            f"(found {source.count(old)})"
        )
    module = types.ModuleType("linexcel.refs")
    module.__file__ = str(path)
    code = compile(source.replace(old, new), str(path) + " (mutant)", "exec")
    # Dataclasses look their module up by name while the class is created.
    saved = sys.modules["linexcel.refs"]
    sys.modules["linexcel.refs"] = module
    try:
        exec(code, module.__dict__)  # noqa: S102 - our own source, one edit
    finally:
        sys.modules["linexcel.refs"] = saved
    return module


def run_mutants(baseline: set[str]) -> int:
    print("Broken variants of src/linexcel/refs.py (in memory, nothing is written):")
    survived = 0
    for label, old, new in MUTANTS:
        try:
            module = mutated_refs(old, new)
        except StaleMutantError as exc:
            print(f"  STALE   {label}\n          {exc}")
            survived += 1
            continue
        try:
            _totals, found = run(Pipeline(module), True, 4, False)
            new_classes = sorted(set(classes(found)) - baseline)
        except Exception as exc:  # noqa: BLE001 - a crash also exposes the break
            new_classes = [f"crash: {type(exc).__name__}: {exc}"]
        if new_classes:
            print(f"  caught  {label}  ({len(new_classes)} new violation classes)")
            print(f"          e.g. {new_classes[0]}")
        else:
            print(f"  MISSED  {label}")
            survived += 1
    if survived:
        print(f"{survived} broken variant(s) went unnoticed: the check is too weak")
        return 1
    print(f"all {len(MUTANTS)} broken variants were caught")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--quick", action="store_true", help="first two windows only")
    parser.add_argument("--mutants", action="store_true")
    parser.add_argument("--show", metavar="CLASS", help="all examples of one class")
    parser.add_argument("--examples", type=int, default=1)
    parser.add_argument("--update-known", action="store_true")
    parser.add_argument("--max-group", type=int, default=4, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        getattr(stream, "reconfigure", lambda **_: None)(encoding="utf-8")

    stale = stale_call_sites()
    if stale:
        print("STALE MODEL: these call-site lines are no longer in the code:")
        for line in stale:
            print(f"  {line}")
        print("The check composes the functions as those lines did. Re-read the")
        print("call sites, update CALL_SITES and Pipeline, then run again.")
        return 1

    known = json.loads(KNOWN.read_text("utf-8")) if KNOWN.exists() else {"classes": {}}
    baseline = set(known["classes"])
    if args.mutants:
        return run_mutants(baseline)

    print(f"linexcel working tree at {ROOT}")
    print(f"formualizer {version('formualizer')} (tokenizer and parser in the loop)")
    started = time.time()
    totals, found = run(Pipeline(), args.quick, args.max_group, True)
    grouped = classes(found)
    print(
        f"checked {totals['formulas']:,} formulas, {totals['groups']:,} group "
        f"shapes, {totals['memberChecks']:,} member checks "
        f"in {time.time() - started:.1f}s"
    )

    if args.update_known:
        KNOWN.write_text(
            json.dumps(
                {
                    "note": known.get("note", ""),
                    "classes": {
                        klass: known["classes"].get(klass, "")
                        or f"{len(items)} counterexamples at recording time"
                        for klass, items in sorted(grouped.items())
                    },
                },
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
        print(f"recorded {len(grouped)} classes in {KNOWN.name}")
        return 0

    if args.show:
        for violation in grouped.get(args.show, []):
            print("\n".join(violation.lines))
        return 0 if args.show in grouped else 1

    new = sorted(set(grouped) - baseline)
    # A class absent from a partial run is not evidence that it was fixed.
    gone = [] if args.quick else sorted(baseline - set(grouped))
    for title, names in (("NEW VIOLATION", new), ("known finding", sorted(baseline))):
        for klass in names:
            items = grouped.get(klass)
            if not items:
                continue
            print(f"\n{title}  {klass}  ({len(items):,} counterexamples)")
            for violation in items[: args.examples]:
                print("  " + "\n  ".join(violation.lines))
    for klass in gone:
        print(f"\nNO LONGER FAILS  {klass}: remove it from {KNOWN.name}")

    print()
    if new:
        print(f"RESULT: FAIL - {len(new)} violation class(es) not in {KNOWN.name}")
        return 1
    if gone:
        print(
            f"RESULT: {len(gone)} recorded finding(s) no longer fail. Update "
            f"{KNOWN.name} (--update-known) and the skill's Findings section."
        )
        return 1
    if grouped:
        print(
            f"RESULT: no new violation; {len(grouped)} known finding class(es) "
            "still fail (see the skill's Findings section)"
        )
        return 3
    print("RESULT: PASS - every group claims every cell its members read")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
