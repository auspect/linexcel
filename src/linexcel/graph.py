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
from typing import Any

from linexcel.external import parse_external_refs
from linexcel.refs import Rect, a1
from linexcel.resolver import _external_name, _ValueResolver
from linexcel.structure import MAX_NODES_PER_SHEET
from linexcel.sweep import FormulaGroup
from linexcel.tables import _enrich_with_table
from linexcel.values import _jsonable

SMALL_RANGE_CELLS = 20_000
MAX_VALUE_SAMPLE = 5


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
