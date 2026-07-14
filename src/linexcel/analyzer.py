"""Builds the lineage graph of an Excel workbook.

Steps:
1. structure (sheets, dimensions, defined names) via openpyxl read-only;
2. formulas + computed values via the Rust engine formualizer;
3. grouping of stretched formulas by R1C1 canonicalization —
   a column of 50,000 copied formulas becomes ONE node;
4. resolution of precedents (cells, ranges, names, other sheets);
5. decomposition of each composite formula into individually evaluated steps
   in a scratch sheet of the engine;
6. lineage of extracted VBA code (oletools).
"""

from __future__ import annotations

import datetime
import io
import itertools
import uuid
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

import formualizer as fz
from openpyxl import load_workbook

from linexcel.refs import (
    Rect,
    num_to_col,
    parse_ref_detailed,
    stretch_ref,
)
from linexcel.rewrite import canonical_r1c1, qualify_sheet
from linexcel.vba import VbaProc, analyze_vba, extract_vba_modules

# Guards to stay responsive on large workbooks.
SCAN_CHUNK_ROWS = 20_000
MAX_CELLS_PER_SHEET = 4_000_000
SMALL_RANGE_CELLS = 20_000
MAX_NODES_PER_SHEET = 400
MAX_STEPS_PER_FORMULA = 48
MAX_SCRATCH_EVALS = 4_000
MAX_VALUE_SAMPLE = 5
MAX_VBA_CODE_CHARS = 6_000
SCRATCH_SHEET = "__lineage_scratch__"


@dataclass
class FormulaGroup:
    """A set of cells on a sheet sharing the same R1C1 formula."""

    sheet: str
    r1c1: str
    cells: list[tuple[int, int]] = field(default_factory=list)
    formulas: dict[tuple[int, int], str] = field(default_factory=dict)

    @property
    def rep(self) -> tuple[int, int]:
        return min(self.cells)

    @property
    def bbox(self) -> tuple[int, int, int, int]:
        rows = [r for r, _ in self.cells]
        cols = [c for _, c in self.cells]
        return min(rows), min(cols), max(rows), max(cols)


class _Budget:
    def __init__(self, limit: int):
        self.left = limit

    def take(self) -> bool:
        if self.left <= 0:
            return False
        self.left -= 1
        return True


def a1(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def analyze_workbook(data: bytes, filename: str = "workbook.xlsx") -> dict[str, Any]:
    """Full analysis: returns the JSON-serializable graph and the engine."""
    warnings: list[str] = []

    # --- 1. structure -----------------------------------------------------
    owb = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    sheet_dims: dict[str, tuple[int, int]] = {}
    for ws in owb.worksheets:
        max_row, max_col = ws.max_row, ws.max_column
        if not max_row or not max_col:
            max_row, max_col = _force_dimensions(ws)
        sheet_dims[ws.title] = (max_row or 1, max_col or 1)
    defined_names = _collect_defined_names(owb)
    owb.close()

    # --- 2. computation engine -------------------------------------------
    engine = fz.Workbook.from_bytes(data)
    engine_sheets = set(engine.sheet_names)
    try:
        engine.evaluate_all()
    except Exception as exc:  # graph remains useful without values
        warnings.append(f"Global evaluation incomplete: {exc}")

    # --- 3. extraction + grouping ----------------------------------------
    groups: dict[tuple[str, str], FormulaGroup] = {}
    cell_owner: dict[str, dict[tuple[int, int], str]] = defaultdict(dict)
    formula_count = 0
    sheet_stats: list[dict[str, Any]] = []

    for sheet, (max_row, max_col) in sheet_dims.items():
        if sheet not in engine_sheets:
            warnings.append(f"Sheet '{sheet}' skipped (not loaded by engine)")
            continue
        n_formulas = 0
        scanned = 0
        fsheet = engine.sheet(sheet)
        for r0 in range(1, max_row + 1, SCAN_CHUNK_ROWS):
            r1 = min(r0 + SCAN_CHUNK_ROWS - 1, max_row)
            if scanned > MAX_CELLS_PER_SHEET:
                warnings.append(
                    f"Sheet '{sheet}' truncated after {scanned:,} cells"
                )
                break
            ra = fz.RangeAddress(sheet, r0, 1, r1, max_col)
            try:
                rows = fsheet.get_formulas(ra)
            except Exception as exc:
                warnings.append(f"Could not read formulas on {sheet}: {exc}")
                break
            scanned += (r1 - r0 + 1) * max_col
            for i, row_vals in enumerate(rows):
                r = r0 + i
                for j, f in enumerate(row_vals):
                    if not f:
                        continue
                    c = j + 1
                    n_formulas += 1
                    key = (sheet, canonical_r1c1(f, r, c))
                    grp = groups.get(key)
                    if grp is None:
                        grp = groups[key] = FormulaGroup(sheet, key[1])
                    grp.cells.append((r, c))
                    # row/col order scan: first cell seen is the representative
                    # (min), keep 3 example formulas
                    if len(grp.formulas) < 3:
                        grp.formulas[(r, c)] = f
        formula_count += n_formulas
        sheet_stats.append(
            {
                "name": sheet,
                "rows": max_row,
                "cols": max_col,
                "formulaCells": n_formulas,
            }
        )

    # --- 4. formula nodes -------------------------------------------------
    nodes: dict[str, dict[str, Any]] = {}
    edges: dict[tuple[str, str, str], dict[str, Any]] = {}

    per_sheet_groups: dict[str, list[FormulaGroup]] = defaultdict(list)
    for grp in groups.values():
        per_sheet_groups[grp.sheet].append(grp)

    kept_groups: list[tuple[str, FormulaGroup]] = []
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
            kept_groups.append((node_id, grp))
            for cell in grp.cells:
                cell_owner[sheet][cell] = node_id
        if dropped:
            n_cells = sum(len(g.cells) for g in dropped)
            misc_id = f"misc:{sheet}"
            nodes[misc_id] = {
                "id": misc_id,
                "kind": "misc",
                "sheet": sheet,
                "label": f"{len(dropped)} other patterns ({n_cells} cells)",
                "count": n_cells,
                "patterns": len(dropped),
            }
            warnings.append(
                f"Sheet '{sheet}': {len(dropped)} formula patterns aggregated "
                f"into a 'misc' node (limit {MAX_NODES_PER_SHEET})"
            )
            for grp in dropped:
                for cell in grp.cells:
                    cell_owner[sheet][cell] = misc_id

    ast_cache: dict[str, Any] = {}
    input_nodes: dict[str, str] = {}  # full A1 key -> node id

    def ensure_input_node(rect: Rect, opaque_label: str | None = None) -> str:
        label = opaque_label or rect.to_a1()
        node_id = input_nodes.get(label)
        if node_id:
            return node_id
        if opaque_label is not None:
            node_id = f"x:{opaque_label}"
            nodes[node_id] = {
                "id": node_id,
                "kind": "opaque",
                "label": opaque_label,
                "sheet": None,
            }
        else:
            node_id = f"i:{label}"
            nodes[node_id] = {
                "id": node_id,
                "kind": "input",
                "label": label,
                "sheet": rect.sheet,
                "addr": label.split("!")[-1],
                "count": rect.ncells,
                "values": _sample_range_values(engine, rect, engine_sheets),
            }
        input_nodes[label] = node_id
        return node_id

    def add_edge(src: str, dst: str, kind: str, approx: bool = False) -> None:
        if src == dst:
            return
        key = (src, dst, kind)
        e = edges.get(key)
        if e is None:
            edges[key] = {
                "id": f"e{len(edges)}",
                "source": src,
                "target": dst,
                "kind": kind,
                "approx": approx,
            }
        elif not approx:
            e["approx"] = False

    def resolve_rect_edges(rect: Rect, target_id: str, kind: str = "dep") -> None:
        """Create precedent → target edges for a referenced range."""
        sheet = rect.sheet
        if sheet not in sheet_dims:
            ensure_input_node(rect, opaque_label=rect.to_a1())
            add_edge(input_nodes[rect.to_a1()], target_id, kind)
            return
        clipped = rect.clipped(*sheet_dims[sheet])
        if clipped is None:
            return
        owners = cell_owner.get(sheet, {})
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
                        add_edge(owner, target_id, kind)
            if has_plain:
                add_edge(ensure_input_node(clipped), target_id, kind)
        else:
            # Huge range: approximate intersection with node bounding boxes.
            for node_id, grp in kept_groups:
                if grp.sheet != sheet:
                    continue
                r1, c1, r2, c2 = grp.bbox
                if clipped.intersects(Rect(sheet, r1, c1, r2, c2)):
                    add_edge(node_id, target_id, kind, approx=True)
            add_edge(ensure_input_node(clipped), target_id, kind, approx=True)

    # defined names -----------------------------------------------------------
    name_nodes: dict[str, str] = {}
    for name, targets in defined_names.items():
        node_id = f"n:{name}"
        name_nodes[name.upper()] = node_id
        val = None
        if targets:
            first = targets[0]
            if first.r1 == first.r2 and first.c1 == first.c2:
                val = _cell_value(
                    engine, first.sheet, first.r1, first.c1, engine_sheets
                )
            else:
                val_samples = _sample_range_values(engine, first, engine_sheets)
                if val_samples:
                    val = val_samples[0]["value"]
        nodes[node_id] = {
            "id": node_id,
            "kind": "name",
            "label": name,
            "sheet": targets[0].sheet if targets else None,
            "targets": [t.to_a1() for t in targets],
            "value": val,
        }
        for rect in targets:
            resolve_rect_edges(rect, node_id, kind="name")

    # formula nodes + edges -------------------------------------------------
    scratch_ready = _ensure_scratch(engine)
    budget = _Budget(MAX_SCRATCH_EVALS)

    for node_id, grp in kept_groups:
        rep_r, rep_c = grp.rep
        formula = grp.formulas.get((rep_r, rep_c)) or next(iter(grp.formulas.values()))
        sheet = grp.sheet
        is_group = len(grp.cells) > 1
        try:
            ast = ast_cache.get(formula)
            if ast is None:
                ast = ast_cache[formula] = fz.parse(
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
                if up in name_nodes:
                    add_edge(name_nodes[up], node_id, "name")
                else:
                    opaque_id = ensure_input_node(
                        Rect(None, 1, 1, 1, 1), opaque_label=ref
                    )
                    add_edge(opaque_id, node_id, "dep")
                continue
            rect = (
                stretch_ref(detail, rep_r, rep_c, (rmin, rmax), (cmin, cmax))
                if is_group
                else detail.rect
            )
            agg_rects.append(rect)

        for rect in _merge_rects(agg_rects):
            resolve_rect_edges(rect, node_id)

        value = _cell_value(engine, sheet, rep_r, rep_c, engine_sheets)
        samples = None
        if is_group:
            samples = []
            for r, c in itertools.islice(sorted(grp.cells), 3):
                samples.append(
                    {
                        "addr": a1(r, c),
                        "value": _cell_value(engine, sheet, r, c, engine_sheets),
                    }
                )

        steps = None
        if ast_dict is not None:
            steps = _decompose(
                ast_dict,
                sheet,
                engine,
                scratch_ready,
                budget,
                engine_sheets,
                defined_names,
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
            "value": value,
            "samples": samples,
            "steps": steps,
        }
        nodes[node_id] = node

    # --- 6. VBA --------------------------------------------------------------
    vba_modules = extract_vba_modules(data, filename)
    vba_procs: list[VbaProc] = analyze_vba(vba_modules) if vba_modules else []
    proc_ids: dict[str, str] = {}
    for proc in vba_procs:
        pid = f"vp:{proc.module}.{proc.name}"
        proc_ids[proc.name] = pid
        nodes[pid] = {
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
    for proc in vba_procs:
        pid = proc_ids[proc.name]
        for callee in proc.calls:
            if callee in proc_ids:
                add_edge(pid, proc_ids[callee], "call")
        for ref in proc.refs:
            detail = parse_ref_detailed(ref.ref, default_sheet=ref.sheet)
            if detail is None or detail.rect.sheet is None:
                opaque_id = ensure_input_node(
                    Rect(None, 1, 1, 1, 1),
                    opaque_label=f"VBA:{ref.sheet or '?'}!{ref.ref}",
                )
                if ref.access == "write":
                    add_edge(pid, opaque_id, "vba-write")
                else:
                    add_edge(opaque_id, pid, "vba-read")
                continue
            if ref.access == "write":
                _resolve_vba_write(
                    detail.rect,
                    pid,
                    sheet_dims,
                    cell_owner,
                    add_edge,
                    ensure_input_node,
                )
            else:
                resolve_rect_edges(detail.rect, pid, kind="vba-read")

    graph = {
        "meta": {
            "filename": filename,
            "analyzedAt": datetime.datetime.now(datetime.UTC).isoformat(),
            "engine": "formualizer (Rust)",
            "warnings": warnings,
            "stats": {
                "sheets": sheet_stats,
                "totalFormulas": formula_count,
                "totalNodes": len(nodes),
                "totalEdges": len(edges),
                "groupedPatterns": sum(1 for _, g in kept_groups if len(g.cells) > 1),
                "vbaModules": len(vba_modules),
                "vbaProcs": len(vba_procs),
                "definedNames": len(defined_names),
            },
        },
        "sheets": list(sheet_dims.keys()),
        "nodes": list(nodes.values()),
        "edges": list(edges.values()),
    }
    return {"graph": graph, "engine": engine, "analysisId": uuid.uuid4().hex[:16]}


# ---------------------------------------------------------------------------
# Composite function decomposition
# ---------------------------------------------------------------------------

_STEP_KINDS = {"Function", "BinaryOp", "UnaryOp"}


def _decompose(
    ast_dict: dict,
    sheet: str,
    engine,
    scratch_ready: bool,
    budget: _Budget,
    engine_sheets: set[str],
    defined_names: dict[str, list[Rect]] | None = None,
) -> dict | None:
    """Step tree: each function / operator becomes an evaluated step."""
    counter = itertools.count()

    def expr_of(node: dict) -> str:
        return _render_expr(node)

    def walk(node: dict, depth: int) -> dict | None:
        ntype = node.get("node_type")
        if ntype not in _STEP_KINDS:
            return None
        if next(counter) >= MAX_STEPS_PER_FORMULA:
            return None
        expr = expr_of(node)
        if ntype == "Function":
            label = node.get("name", "?")
            children_ast = node.get("args", [])
        elif ntype == "BinaryOp":
            label = node.get("operator", "?")
            children_ast = [node.get("left"), node.get("right")]
        else:
            label = node.get("operator", "?")
            children_ast = [node.get("operand") or node.get("expr")]
        children_ast = [c for c in children_ast if c]

        inputs = []
        children = []
        for child in children_ast:
            sub = walk(child, depth + 1)
            if sub is not None:
                children.append(sub)
            else:
                ctype = child.get("node_type")
                if ctype == "Reference":
                    ref = child.get("reference", "?")
                    inputs.append(
                        {
                            "ref": ref,
                            "value": _ref_preview(
                                engine, ref, sheet, engine_sheets, defined_names
                            ),
                        }
                    )
                elif ctype == "Literal":
                    inputs.append({"literal": child.get("value")})

        value, evaluated = None, False
        if scratch_ready and budget.take():
            value, evaluated = _scratch_eval(engine, expr, sheet)
        return {
            "kind": ntype,
            "label": label,
            "expr": expr,
            "value": value,
            "evaluated": evaluated,
            "inputs": inputs,
            "children": children,
        }

    return walk(ast_dict, 0)


def _render_expr(node: dict) -> str:
    """Reconstruct the expression of an AST subtree (readable form)."""
    ntype = node.get("node_type")
    if ntype == "Function":
        args = ", ".join(_render_expr(a) for a in node.get("args", []))
        return f"{node.get('name', '?')}({args})"
    if ntype == "BinaryOp":
        return (
            f"{_render_expr(node.get('left', {}))} {node.get('operator', '?')} "
            f"{_render_expr(node.get('right', {}))}"
        )
    if ntype == "UnaryOp":
        operand = node.get("operand") or node.get("expr") or {}
        return f"{node.get('operator', '?')}{_render_expr(operand)}"
    if ntype == "Reference":
        return str(node.get("reference", "?"))
    if ntype == "Literal":
        v = node.get("value")
        if isinstance(v, str):
            return '"' + v.replace('"', '""') + '"'
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    if ntype == "Array":
        return "{...}"
    if ntype == "Paren":
        inner = node.get("expr") or node.get("inner") or {}
        return f"({_render_expr(inner)})"
    return "?"


def _scratch_eval(engine, expr: str, sheet: str) -> tuple[Any, bool]:
    try:
        qualified = qualify_sheet(expr, sheet)
        engine.set_formula(SCRATCH_SHEET, 1, 1, qualified)
        value = engine.evaluate_cell(SCRATCH_SHEET, 1, 1)
        return _jsonable(value), True
    except Exception:
        return None, False


def _ensure_scratch(engine) -> bool:
    try:
        engine.add_sheet(SCRATCH_SHEET)
        return True
    except Exception:
        return SCRATCH_SHEET in set(engine.sheet_names)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _force_dimensions(ws) -> tuple[int, int]:
    max_row = max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
        if max_row > 200_000:
            break
    return max_row or 1, max_col or 1


def _collect_defined_names(owb) -> dict[str, list[Rect]]:
    out: dict[str, list[Rect]] = {}
    try:
        items = owb.defined_names.items()
    except Exception:
        return out
    for name, dn in items:
        if name.startswith("_xlnm."):
            continue
        rects: list[Rect] = []
        try:
            for sheet, coord in dn.destinations:
                detail = parse_ref_detailed(coord, default_sheet=sheet)
                if detail is not None:
                    rect = detail.rect
                    rects.append(
                        Rect(sheet or rect.sheet, rect.r1, rect.c1, rect.r2, rect.c2)
                    )
        except Exception:
            continue
        if rects:
            out[name] = rects
    return out


def _collect_ref_strings(ast_dict: dict) -> list[str]:
    refs: list[str] = []

    def walk(node) -> None:
        if isinstance(node, dict):
            if node.get("node_type") == "Reference":
                ref = node.get("reference")
                if ref:
                    refs.append(str(ref))
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    walk(ast_dict)
    # dedupe preserving order
    seen: set[str] = set()
    out = []
    for r in refs:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out


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


def _cell_value(engine, sheet: str, row: int, col: int, engine_sheets: set[str]):
    if sheet not in engine_sheets:
        return None
    try:
        return _jsonable(engine.get_value(sheet, row, col))
    except Exception:
        try:
            return _jsonable(engine.evaluate_cell(sheet, row, col))
        except Exception:
            return None


def _sample_range_values(engine, rect: Rect, engine_sheets: set[str]) -> list:
    if rect.sheet is None or rect.sheet not in engine_sheets:
        return []
    out = []
    try:
        for r in range(rect.r1, min(rect.r2, rect.r1 + MAX_VALUE_SAMPLE - 1) + 1):
            for c in range(rect.c1, min(rect.c2, rect.c1 + MAX_VALUE_SAMPLE - 1) + 1):
                if len(out) >= MAX_VALUE_SAMPLE:
                    return out
                out.append(
                    {
                        "addr": a1(r, c),
                        "value": _jsonable(engine.get_value(rect.sheet, r, c)),
                    }
                )
    except Exception:
        pass
    return out


def _ref_preview(
    engine,
    ref: str,
    sheet: str,
    engine_sheets: set[str],
    defined_names: dict[str, list[Rect]] | None = None,
):
    detail = parse_ref_detailed(ref, default_sheet=sheet)
    if detail is None:
        # may be a defined name: show the value of its target
        if defined_names:
            for name, rects in defined_names.items():
                if name.upper() == ref.upper() and rects:
                    rect = rects[0]
                    if rect.ncells == 1:
                        return _cell_value(
                            engine, rect.sheet or sheet, rect.r1, rect.c1, engine_sheets
                        )
                    return {"range": rect.to_a1(), "n": rect.ncells}
        return None
    rect = detail.rect
    if rect.ncells == 1:
        return _cell_value(engine, rect.sheet or sheet, rect.r1, rect.c1, engine_sheets)
    return {"range": rect.to_a1(), "n": rect.ncells}


def _resolve_vba_write(
    rect: Rect, pid: str, sheet_dims, cell_owner, add_edge, ensure_input_node
) -> None:
    """A VBA write feeds the target cells: edge proc → target."""
    sheet = rect.sheet
    if sheet not in sheet_dims:
        opaque = ensure_input_node(rect, opaque_label=rect.to_a1())
        add_edge(pid, opaque, "vba-write")
        return
    clipped = rect.clipped(*sheet_dims[sheet]) or rect
    owners = cell_owner.get(sheet, {})
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
                    add_edge(pid, owner, "vba-write")
    else:
        has_plain = True
    if has_plain:
        add_edge(pid, ensure_input_node(clipped), "vba-write")


def _jsonable(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        is_nan_or_inf = value != value or value in (float("inf"), float("-inf"))
        if isinstance(value, float) and is_nan_or_inf:
            return str(value)
        return value
    return str(value)
