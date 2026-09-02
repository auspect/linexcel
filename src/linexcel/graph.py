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

from linexcel.refs import a1
from linexcel.resolver import _ValueResolver
from linexcel.structure import MAX_NODES_PER_SHEET
from linexcel.sweep import FormulaGroup


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
