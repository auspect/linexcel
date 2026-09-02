"""Builds the graph's nodes and edges.

Extracted mechanically from analyzer.py: the eight closures of
``analyze_workbook`` that mutually capture ``nodes``/``edges``/``cell_owner``/
etc. become methods of ``GraphBuilder``, one state object that owns those
dicts instead of a function-local tangle every closure reaches into.

Built up in micro-commits, in this order: node selection, edge/node
infrastructure (``add_edge``, ``ensure_opaque_node``, ``ensure_input_node``,
``resolve_rect_edges``), defined names, formula nodes, VBA, Power Query.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

import formualizer as fz

from linexcel.decompose import _collect_step_exprs, _decompose
from linexcel.external import macro_files, parse_external_refs
from linexcel.loader import _stepped
from linexcel.refs import Rect, a1, parse_ref_detailed, stretch_ref
from linexcel.resolver import _collect_ref_strings, _external_name, _ValueResolver
from linexcel.structure import MAX_NODES_PER_SHEET
from linexcel.sweep import FormulaGroup
from linexcel.tables import _enrich_with_table
from linexcel.values import _jsonable
from linexcel.vba import VbaProc, analyze_vba, extract_vba_modules

SMALL_RANGE_CELLS = 20_000
MAX_VALUE_SAMPLE = 5
MAX_VBA_CODE_CHARS = 6_000


def _merge_rects(rects: list[Rect]) -> list[Rect]:
    seen: set[tuple] = set()
    out = []
    for r in rects:
        key = (r.sheet, r.r1, r.c1, r.r2, r.c2)
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def _bbox_a1(grp: FormulaGroup) -> str:
    r1, c1, r2, c2 = grp.bbox
    if (r1, c1) == (r2, c2):
        return a1(r1, c1)
    return f"{a1(r1, c1)}:{a1(r2, c2)}"


def _spread_cells(cells: list[tuple[int, int]], n: int) -> list[tuple[int, int]]:
    """``n`` cells spread evenly over a group, in reading order.

    The head of a stretch is made of near-identical neighbours: sampling
    ``B2, C2, B3`` out of 400,000 cells says nothing about the far end, which
    is exactly where a pattern that broke — a hard-coded cell dropped into the
    middle of a column — shows up. First and last are always included.
    """
    ordered = sorted(cells)
    if n < 2:
        return ordered[: max(n, 0)]
    if len(ordered) <= n:
        return ordered
    last = len(ordered) - 1
    picks = sorted({round(i * last / (n - 1)) for i in range(n)})
    return [ordered[i] for i in picks]


def _resolve_call(
    callee: str,
    caller_module: str,
    proc_ids: dict[str, str],
    procs_by_name: dict[str, list[str]],
) -> str | None:
    """Resolve a VBA call to a procedure node id.

    A callee already qualified with its module (``Module1.Taux``) resolves
    directly. VBA looks an unqualified name up in the calling module first,
    then in the other modules; a name declared by several other modules stays
    unresolved rather than pointing at an arbitrary one. All lookups are
    case-insensitive, as the language is.
    """
    if "." in callee:
        return proc_ids.get(callee.lower())
    same_module = f"{caller_module}.{callee}".lower()
    if same_module in proc_ids:
        return proc_ids[same_module]
    candidates = procs_by_name.get(callee.lower(), [])
    if len(candidates) == 1:
        return proc_ids[candidates[0]]
    return None


def _sample_range_values(resolver: _ValueResolver, rect: Rect) -> list:
    if rect.sheet is None:
        return []
    out = []
    for r in range(rect.r1, min(rect.r2, rect.r1 + MAX_VALUE_SAMPLE - 1) + 1):
        for c in range(rect.c1, min(rect.c2, rect.c1 + MAX_VALUE_SAMPLE - 1) + 1):
            if len(out) >= MAX_VALUE_SAMPLE:
                return out
            value, source, date_text = resolver.value(rect.sheet, r, c)
            out.append(
                {
                    "addr": a1(r, c),
                    "value": value,
                    "source": source,
                    "date": date_text,
                }
            )
    return out


class GraphBuilder:
    """Owns the graph's nodes/edges; the analyzer's closures become methods."""

    def __init__(
        self,
        resolver: _ValueResolver,
        sheet_dims: dict[str, tuple[int, int]],
        table_index: dict[str, list[dict[str, Any]]],
        defined_names: dict[str, list],
        warnings: list[str],
        reporter: Any,
    ) -> None:
        self.resolver = resolver
        self.sheet_dims = sheet_dims
        self.table_index = table_index
        self.defined_names = defined_names
        self.warnings = warnings
        self.reporter = reporter

        self.nodes: dict[str, dict[str, Any]] = {}
        self.edges: dict[tuple[str, str, str], dict[str, Any]] = {}
        self.input_nodes: dict[str, str] = {}  # full A1 key -> node id
        self.name_nodes: dict[str, str] = {}
        self.cell_owner: dict[str, dict[tuple[int, int], str]] = defaultdict(dict)
        self.ast_cache: dict[str, Any] = {}
        self.kept_groups: list[tuple[str, FormulaGroup]] = []

    def select_nodes(self, groups: dict[tuple[str, str], FormulaGroup]) -> None:
        """Cap nodes per sheet; fold what's dropped into one 'misc' node."""
        per_sheet_groups: dict[str, list[FormulaGroup]] = defaultdict(list)
        for grp in groups.values():
            per_sheet_groups[grp.sheet].append(grp)

        for sheet, sheet_groups in per_sheet_groups.items():
            sheet_groups.sort(key=lambda g: (-len(g.cells), g.rep))
            kept = sheet_groups[:MAX_NODES_PER_SHEET]
            dropped = sheet_groups[MAX_NODES_PER_SHEET:]
            for grp in kept:
                rep_r, rep_c = grp.rep
                if len(grp.cells) == 1:
                    node_id = f"c:{sheet}!{a1(rep_r, rep_c)}"
                else:
                    node_id = f"g:{sheet}!{a1(rep_r, rep_c)}#{len(grp.cells)}"
                self.kept_groups.append((node_id, grp))
                for cell in grp.cells:
                    self.cell_owner[sheet][cell] = node_id
            if dropped:
                n_cells = sum(len(g.cells) for g in dropped)
                misc_id = f"misc:{sheet}"
                self.nodes[misc_id] = {
                    "id": misc_id,
                    "kind": "misc",
                    "sheet": sheet,
                    "label": f"{len(dropped)} other patterns ({n_cells} cells)",
                    "count": n_cells,
                    "patterns": len(dropped),
                }
                self.warnings.append(
                    f"Sheet '{sheet}': {len(dropped)} formula patterns aggregated "
                    f"into a 'misc' node (limit {MAX_NODES_PER_SHEET})"
                )
                for grp in dropped:
                    for cell in grp.cells:
                        self.cell_owner[sheet][cell] = misc_id

    def ensure_opaque_node(self, ref: str) -> str:
        """A reference the graph cannot follow, named as precisely as it can be.

        An external reference is one of those, but it is not opaque to the
        *reader*: the file says which workbook it points at, and the node
        carries that — with the value, when the workbook could be read.
        """
        external = parse_external_refs(ref)
        if not external:
            return self.ensure_input_node(Rect(None, 1, 1, 1, 1), opaque_label=ref)
        first = external[0]
        books = self.resolver.external_books(ref)
        name = books[0]["name"] if books else _external_name(first)
        label = f"[{name}]{first.sheet}!{first.cell}" if first.sheet else f"[{name}]"
        node_id = self.input_nodes.get(label)
        if node_id:
            return node_id
        node_id = f"x:{label}"
        node: dict[str, Any] = {
            "id": node_id,
            "kind": "opaque",
            "label": label,
            "sheet": None,
            "ref": ref,
            "externalBooks": books,
        }
        value, source = self.resolver.external_value(first)
        if source is not None:
            node["value"] = _jsonable(value)
            node["valueSource"] = source
        self.nodes[node_id] = node
        self.input_nodes[label] = node_id
        return node_id

    def ensure_input_node(self, rect: Rect, opaque_label: str | None = None) -> str:
        label = opaque_label or rect.to_a1()
        node_id = self.input_nodes.get(label)
        if node_id:
            return node_id
        if opaque_label is not None:
            node_id = f"x:{opaque_label}"
            self.nodes[node_id] = {
                "id": node_id,
                "kind": "opaque",
                "label": opaque_label,
                "sheet": None,
            }
        else:
            node_id = f"i:{label}"
            node: dict[str, Any] = {
                "id": node_id,
                "kind": "input",
                "label": label,
                "sheet": rect.sheet,
                "addr": label.split("!")[-1],
                "count": rect.ncells,
                "values": _sample_range_values(self.resolver, rect),
            }
            if rect.ncells == 1 and rect.sheet is not None:
                node.update(self.resolver.describe(rect.sheet, rect.r1, rect.c1))
                _enrich_with_table(
                    node, self.table_index, rect.sheet, rect.r1, rect.c1
                )
            self.nodes[node_id] = node
        self.input_nodes[label] = node_id
        return node_id

    def add_edge(self, src: str, dst: str, kind: str, approx: bool = False) -> None:
        if src == dst:
            return
        key = (src, dst, kind)
        e = self.edges.get(key)
        if e is None:
            self.edges[key] = {
                "id": f"e{len(self.edges)}",
                "source": src,
                "target": dst,
                "kind": kind,
                "approx": approx,
            }
        elif not approx:
            e["approx"] = False

    def resolve_rect_edges(self, rect: Rect, target_id: str, kind: str = "dep") -> None:
        """Create precedent → target edges for a referenced range."""
        sheet = rect.sheet
        if sheet not in self.sheet_dims:
            # A sheet this file does not have: another workbook, most often —
            # ``'[1]Annual'!B4`` parses as a sheet name, brackets and all.
            self.add_edge(self.ensure_opaque_node(rect.to_a1()), target_id, kind)
            return
        clipped = rect.clipped(*self.sheet_dims[sheet])
        if clipped is None:
            return
        owners = self.cell_owner.get(sheet, {})
        if clipped.ncells <= SMALL_RANGE_CELLS:
            seen: set[str] = set()
            has_plain = False
            for r in range(clipped.r1, clipped.r2 + 1):
                for c in range(clipped.c1, clipped.c2 + 1):
                    owner = owners.get((r, c))
                    if owner is None:
                        has_plain = True
                    elif owner not in seen:
                        seen.add(owner)
                        self.add_edge(owner, target_id, kind)
            if has_plain:
                self.add_edge(self.ensure_input_node(clipped), target_id, kind)
        else:
            # Huge range: approximate intersection with node bounding boxes.
            for node_id, grp in self.kept_groups:
                if grp.sheet != sheet:
                    continue
                r1, c1, r2, c2 = grp.bbox
                if clipped.intersects(Rect(sheet, r1, c1, r2, c2)):
                    self.add_edge(node_id, target_id, kind, approx=True)
            self.add_edge(
                self.ensure_input_node(clipped), target_id, kind, approx=True
            )

    def build_names(self) -> None:
        """A node per defined name, wired to whatever it points at."""
        for name, targets in self.defined_names.items():
            node_id = f"n:{name}"
            self.name_nodes[name.upper()] = node_id
            value_fields: dict[str, Any] = {"value": None}
            if targets:
                first = targets[0]
                if (
                    first.sheet is not None
                    and first.r1 == first.r2
                    and first.c1 == first.c2
                ):
                    value_fields = self.resolver.describe(
                        first.sheet, first.r1, first.c1
                    )
                else:
                    val_samples = _sample_range_values(self.resolver, first)
                    if val_samples:
                        value_fields = {"value": val_samples[0]["value"]}
            self.nodes[node_id] = {
                "id": node_id,
                "kind": "name",
                "label": name,
                "sheet": targets[0].sheet if targets else None,
                "targets": [t.to_a1() for t in targets],
                **value_fields,
            }
            for rect in targets:
                self.resolve_rect_edges(rect, node_id, kind="name")

    def build_formula_nodes(self) -> None:
        """A node per kept formula pattern, wired to its precedents.

        Reported one node at a time: this is where a dense workbook spends
        its minutes, and where a run that looks stuck actually is. A phase
        that prints only when it ends cannot tell anyone that.
        """
        for node_id, grp in _stepped(
            self.reporter.phase("nodes+edges", total=len(self.kept_groups)),
            self.kept_groups,
            "building",
        ):
            rep_r, rep_c = grp.rep
            formula = grp.formulas.get((rep_r, rep_c)) or next(
                iter(grp.formulas.values())
            )
            sheet = grp.sheet
            is_group = len(grp.cells) > 1
            try:
                ast = self.ast_cache.get(formula)
                if ast is None:
                    ast = self.ast_cache[formula] = fz.parse(
                        formula if formula.startswith("=") else "=" + formula
                    )
                ast_dict = ast.to_dict()
            except Exception:
                ast, ast_dict = None, None

            refs = _collect_ref_strings(ast_dict) if ast_dict else []
            rmin, cmin, rmax, cmax = grp.bbox
            agg_rects: list[Rect] = []
            for ref in refs:
                detail = parse_ref_detailed(ref, default_sheet=sheet)
                if detail is None:
                    up = ref.upper()
                    if up in self.name_nodes:
                        self.add_edge(self.name_nodes[up], node_id, "name")
                    else:
                        self.add_edge(self.ensure_opaque_node(ref), node_id, "dep")
                    continue
                rect = (
                    stretch_ref(detail, rep_r, rep_c, (rmin, rmax), (cmin, cmax))
                    if is_group
                    else detail.rect
                )
                agg_rects.append(rect)

            for rect in _merge_rects(agg_rects):
                self.resolve_rect_edges(rect, node_id)

            value_fields = self.resolver.describe(sheet, rep_r, rep_c, formula)
            samples = None
            if is_group:
                samples = [
                    {
                        "addr": a1(r, c),
                        **self.resolver.describe(sheet, r, c, grp.formulas.get((r, c))),
                    }
                    for r, c in _spread_cells(grp.cells, MAX_VALUE_SAMPLE)
                ]

            steps = None
            # A volatile cell is shown as *not* recalculated, so decomposing it
            # would contradict its own card: every step under `=TODAY()+7`
            # would carry a figure computed from today's clock.
            if ast_dict is not None and value_fields.get("valueSource") != "volatile":
                # The root step is the formula itself: when the engine computed
                # the cell, its value is that step's value and needs no
                # scratch pass.
                root_value = (
                    value_fields["value"]
                    if value_fields.get("valueSource") == "engine"
                    else None
                )
                step_exprs = _collect_step_exprs(
                    ast_dict, skip_root=root_value is not None
                )
                if step_exprs:
                    self.resolver.preload_steps(step_exprs, sheet)
                steps = _decompose(
                    ast_dict, sheet, self.resolver, self.defined_names,
                    root_value=root_value,
                )

            node: dict[str, Any] = {
                "id": node_id,
                "kind": "group" if is_group else "cell",
                "sheet": sheet,
                "addr": a1(rep_r, rep_c),
                "label": (
                    f"{sheet}!{a1(rep_r, rep_c)}"
                    + (f" x{len(grp.cells)}" if is_group else "")
                ),
                "formula": formula if formula.startswith("=") else "=" + formula,
                "r1c1": grp.r1c1,
                "count": len(grp.cells),
                "bbox": _bbox_a1(grp),
                **value_fields,
                "samples": samples,
                "steps": steps,
            }
            books = self.resolver.external_books(formula)
            if books:
                node["externalBooks"] = books
            _enrich_with_table(node, self.table_index, sheet, rep_r, rep_c)
            self.nodes[node_id] = node

    def build_vba(
        self, data: bytes, filename: str, refs_dir: str | Path | None
    ) -> None:
        """Nodes for every VBA procedure, wired by call and by cell access."""
        self.vba_modules = extract_vba_modules(data, filename, self.warnings)
        self.vba_procs: list[VbaProc] = (
            analyze_vba(self.vba_modules) if self.vba_modules else []
        )
        # Code a workbook calls often does not live in it: an .xlam add-in
        # holds the functions, and the workbook only names them. Given the
        # folder, that code is read too, and each module says which file it
        # came from.
        if refs_dir is not None:
            for addin in macro_files(Path(refs_dir)):
                extra = extract_vba_modules(
                    addin.read_bytes(), addin.name, self.warnings
                )
                if not extra:
                    continue
                origin = {f"{addin.name}:{name}": code for name, code in extra.items()}
                self.vba_modules.update(origin)
                self.vba_procs.extend(analyze_vba(origin))
        # Node ids keep the declared spelling, but both lookups are keyed on
        # the lowercased name: VBA is case-insensitive, so Module1.Taux and
        # module1.TAUX designate the same procedure. proc_ids resolves a
        # qualified name, procs_by_name the unqualified ones _find_calls
        # reports.
        proc_ids: dict[str, str] = {}
        procs_by_name: dict[str, list[str]] = defaultdict(list)
        for proc in self.vba_procs:
            qualified = f"{proc.module}.{proc.name}"
            pid = f"vp:{qualified}"
            proc_ids[qualified.lower()] = pid
            procs_by_name[proc.name.lower()].append(qualified.lower())
            self.nodes[pid] = {
                "id": pid,
                "kind": "vba",
                "label": f"{proc.module}.{proc.name}",
                "sheet": None,
                "module": proc.module,
                "proc": proc.name,
                "procKind": proc.kind,
                "lines": [proc.line_start, proc.line_end],
                "code": proc.code[:MAX_VBA_CODE_CHARS],
            }
        for proc in self.vba_procs:
            pid = proc_ids[f"{proc.module}.{proc.name}".lower()]
            for callee in proc.calls:
                target = _resolve_call(callee, proc.module, proc_ids, procs_by_name)
                if target is not None:
                    self.add_edge(pid, target, "call")
            for ref in proc.refs:
                detail = parse_ref_detailed(ref.ref, default_sheet=ref.sheet)
                if detail is None or detail.rect.sheet is None:
                    opaque_id = self.ensure_input_node(
                        Rect(None, 1, 1, 1, 1),
                        opaque_label=f"VBA:{ref.sheet or '?'}!{ref.ref}",
                    )
                    if ref.access == "write":
                        self.add_edge(pid, opaque_id, "vba-write")
                    else:
                        self.add_edge(opaque_id, pid, "vba-read")
                    continue
                if ref.access == "write":
                    self._resolve_vba_write(detail.rect, pid)
                else:
                    self.resolve_rect_edges(detail.rect, pid, kind="vba-read")

    def _resolve_vba_write(self, rect: Rect, pid: str) -> None:
        """A VBA write feeds the target cells: edge proc → target."""
        sheet = rect.sheet
        if sheet not in self.sheet_dims:
            opaque = self.ensure_input_node(rect, opaque_label=rect.to_a1())
            self.add_edge(pid, opaque, "vba-write")
            return
        clipped = rect.clipped(*self.sheet_dims[sheet]) or rect
        owners = self.cell_owner.get(sheet, {})
        seen: set[str] = set()
        has_plain = False
        if clipped.ncells <= SMALL_RANGE_CELLS:
            for r in range(clipped.r1, clipped.r2 + 1):
                for c in range(clipped.c1, clipped.c2 + 1):
                    owner = owners.get((r, c))
                    if owner is None:
                        has_plain = True
                    elif owner not in seen:
                        seen.add(owner)
                        self.add_edge(pid, owner, "vba-write")
        else:
            has_plain = True
        if has_plain:
            self.add_edge(pid, self.ensure_input_node(clipped), "vba-write")
