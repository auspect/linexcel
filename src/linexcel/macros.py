"""Static VBA add-in reading, with immutable file and decoded-source fingerprints.

No Excel automation is used. Calls and literal range accesses are heuristic
source observations, never evidence that a procedure ran or a range was resolved.
Legacy XLA worksheet formulas/values are outside this reader's scope.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import uuid
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any

from linexcel.external import macro_files
from linexcel.vba import analyze_vba, extract_vba_modules

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAX_CODE_CHARACTERS = 6000


def valid_excel_header(header: bytes, filename: str) -> bool:
    """Accept OOXML, or an OLE compound document explicitly named as an XLA."""
    if filename.lower().endswith(".xla"):
        return header.startswith(OLE_MAGIC)
    return header.startswith(b"PK")


def read_macro_project(data: bytes, filename: str) -> dict[str, Any]:
    """Read complete decoded modules/procedures without running their code.

    A hash of decoded text is separate from the hash of the original binary.
    Locked, corrupt and partially readable projects retain extraction warnings.
    Neither a missing VBA procedure nor a heuristic call graph proves that a
    workbook contains no Excel 4.0 macros, compiled code or dynamic calls.
    """
    warnings: list[str] = []
    origins: dict[str, dict] = {}
    modules = extract_vba_modules(data, filename, warnings, provenance=origins)
    procedures = []
    for proc in analyze_vba(modules):
        procedures.append(
            {
                "module": proc.module,
                "name": proc.name,
                "kind": proc.kind,
                "lines": [proc.line_start, proc.line_end],
                "code": proc.code,
                "calls": proc.calls,
                "literalReferences": [
                    {
                        "sheet": ref.sheet,
                        "workbook": ref.workbook,
                        "reference": ref.ref,
                        "access": ref.access,
                        "line": ref.line,
                        "resolution": "static_literal_unresolved_workbook_context",
                    }
                    for ref in proc.refs
                ],
                "provenance": {
                    **origins[proc.module],
                    "procedureSha256": hashlib.sha256(proc.code.encode()).hexdigest(),
                },
            }
        )
    return {
        "filename": filename,
        "sha256": hashlib.sha256(data).hexdigest(),
        "mode": "static_vba_source",
        "execution": "not_executed",
        "status": "partial" if warnings else "read",
        "modules": modules,
        "moduleProvenance": origins,
        "procedures": procedures,
        "warnings": warnings,
        "limitations": [
            "VBA source extraction and heuristic calls/range accesses only.",
            "No macro execution, UDF calculation, p-code audit "
            "or Excel 4.0 macro lineage.",
            "Dynamic calls and workbook/sheet aliases can remain unresolved.",
            "Decoded source text is not a byte-for-byte copy of the binary streams.",
        ],
    }


def _macro_graph(
    data: bytes | None, filename: str, refs_dir=None, *, source_hash=None
) -> dict:
    from linexcel.graph import _resolve_call

    main = (
        read_macro_project(data, filename)
        if data is not None
        else {
            "filename": filename,
            "sha256": source_hash,
            "modules": {},
            "procedures": [],
            "warnings": [],
            "status": "read",
            "limitations": [
                "Static VBA only; no execution or complete dynamic-call lineage."
            ],
        }
    )
    projects = [(main, "")]
    if refs_dir is not None:
        for path in macro_files(Path(refs_dir)):
            extra = read_macro_project(path.read_bytes(), path.name)
            if extra["sha256"] == main["sha256"] and path.name == filename:
                continue
            projects.append((extra, f"{path.name}:"))
    nodes: dict[str, dict] = {}
    edges: dict[tuple[str, str, str], dict] = {}
    warnings = []
    module_count = procedure_count = 0
    for project, prefix in projects:
        warnings.extend(f"{project['filename']}: {w}" for w in project["warnings"])
        module_count += len(project["modules"])
        procedure_count += len(project["procedures"])
        ids = {}
        by_name = defaultdict(list)
        declarations = defaultdict(list)
        for proc in project["procedures"]:
            declarations[f"{prefix}{proc['module']}.{proc['name']}".lower()].append(
                proc
            )
        entries = []
        for proc in project["procedures"]:
            qualified = f"{prefix}{proc['module']}.{proc['name']}"
            nid = f"vp:{qualified}"
            if len(declarations[qualified.lower()]) > 1:
                nid += f"@{proc['kind'].replace(' ', '')}:{proc['lines'][0]}"
                ids[qualified.lower()] = None
            else:
                ids[qualified.lower()] = nid
            by_name[proc["name"].lower()].append(qualified.lower())
            entries.append((proc, nid))
            nodes[nid] = {
                "id": nid,
                "kind": "vba",
                "sheet": None,
                "name": qualified,
                "label": qualified,
                "module": f"{prefix}{proc['module']}",
                "proc": proc["name"],
                "procKind": proc["kind"],
                "lines": proc["lines"],
                "code": proc["code"][:MAX_CODE_CHARACTERS],
                "vbaProvenance": {
                    **proc["provenance"],
                    "codeTruncated": len(proc["code"]) > MAX_CODE_CHARACTERS,
                    "totalCodeCharacters": len(proc["code"]),
                },
                "value": None,
                "provenance": "static_vba_source",
                "diagnostics": [],
                "dependencies": [],
            }

        def edge(source, target, kind):
            edges[source, target, kind] = {
                "source": source,
                "target": target,
                "kind": kind,
                "scope": "static_vba_heuristic",
                "approx": True,
            }

        for proc, nid in entries:
            module = f"{prefix}{proc['module']}"
            for callee in proc["calls"]:
                # A qualified reference belongs to this project's module space.
                callee = f"{prefix}{callee}" if "." in callee else callee
                target = _resolve_call(callee, module, ids, by_name)
                if target is not None:
                    edge(nid, target, "call")
                else:
                    nodes[nid]["diagnostics"].append(
                        f"Unresolved or ambiguous static call: {callee}"
                    )
            for ref in proc["literalReferences"]:
                text = (
                    f"{project['filename']}:"
                    f"{ref['workbook'] or 'implicit_active_workbook'}:"
                    f"{ref['sheet'] or '?'}!{ref['reference']}"
                )
                rid = "vba-ref:" + hashlib.sha256(text.encode()).hexdigest()
                nodes.setdefault(
                    rid,
                    {
                        "id": rid,
                        "kind": "opaque",
                        "sheet": None,
                        "label": text,
                        "reference": text,
                        "referenceType": "vba_literal_unresolved_context",
                        "provenance": "static_vba_literal",
                        "value": None,
                        "dependencies": [],
                        "diagnostics": [],
                    },
                )
                if ref["access"] == "write":
                    edge(nid, rid, "vba-write")
                else:
                    edge(rid, nid, "vba-read")
    for relation in edges.values():
        nodes[relation["target"]]["dependencies"].append(relation["source"])
    return {
        "nodes": list(nodes.values()),
        "edges": list(edges.values()),
        "vba": {
            "modules": module_count,
            "procedures": procedure_count,
            "files": [
                {
                    "filename": p["filename"],
                    "sha256": p["sha256"],
                    "status": p["status"],
                }
                for p, _ in projects
            ],
            "warnings": warnings,
            "execution": "not_executed",
            "limitations": main["limitations"],
        },
    }


def analyze_macro_addin(data, filename, *, refs_dir=None, targets=None):
    """Add-in report: static macro graph, explicitly no worksheet analysis."""
    if targets:
        raise ValueError("Cell targets are unavailable for a macro-only add-in report")
    graph = _macro_graph(data, filename, refs_dir)
    macro = graph.pop("vba")
    warnings = list(macro["warnings"])
    scope = (
        "Legacy XLA" if filename.lower().endswith(".xla") else "Add-in source report"
    )
    warnings.append(
        f"{scope}: static VBA only; worksheet cells, cached values "
        "and UDF execution were not analyzed."
    )
    graph["sheets"] = []
    graph["meta"] = {
        "filename": filename,
        "sourceSha256": hashlib.sha256(data).hexdigest(),
        "analyzedAt": dt.datetime.now(dt.UTC).isoformat(),
        "engine": "oletools.olevba (static source)",
        "macroOnly": True,
        "warnings": warnings,
        "vba": macro,
        "analysisCoverage": {
            "requested": "macro_source_only",
            "extractedFormulaCells": None,
            "worksheetCells": "not_analyzed",
            "evaluation": "not_requested",
            "dependencyCompleteness": "not_certified",
            "omissions": [
                {"phase": "worksheet lineage", "reason": "addin_macro_source_only"}
            ],
        },
        "stats": {
            "sheets": [],
            "totalFormulas": 0,
            "totalNodes": len(graph["nodes"]),
            "totalEdges": len(graph["edges"]),
            "groupedPatterns": 0,
            "vbaModules": macro["modules"],
            "vbaProcs": macro["procedures"],
            "definedNames": 0,
            "tables": 0,
            "externalWorkbooks": 0,
            "externalWorkbooksRead": 0,
            "queries": 0,
            "queriesLoaded": 0,
        },
    }
    return {"graph": graph, "engine": None, "analysisId": uuid.uuid4().hex}


def index_vba(store, path: Path, refs_dir=None, *, filename=None) -> dict:
    """Insert static macro source into an already-created sparse disk index."""
    with path.open("rb") as source:
        ole = source.read(8).startswith(OLE_MAGIC)
    has_vba = ole
    if not ole:
        with zipfile.ZipFile(path) as package:
            has_vba = any(
                name.lower().endswith("vbaproject.bin") for name in package.namelist()
            )
    # Do not materialize a large data-only workbook to discover it has no VBA.
    if not has_vba and (refs_dir is None or not macro_files(Path(refs_dir))):
        return {
            "modules": 0,
            "procedures": 0,
            "files": [],
            "warnings": [],
            "execution": "not_executed",
        }
    with path.open("rb") as source:
        source_hash = hashlib.file_digest(source, "sha256").hexdigest()
    graph = _macro_graph(
        path.read_bytes() if has_vba else None,
        filename or path.name,
        refs_dir,
        source_hash=source_hash,
    )
    for node in graph["nodes"]:
        store.put(node)
    for edge in graph["edges"]:
        store.db.execute(
            "INSERT OR IGNORE INTO edges VALUES (?,?)", (edge["source"], edge["target"])
        )
    return graph["vba"]


def static_edge_kind(source_kind, target_kind):
    """Recover the recorded relation type in the lazy graph's two-column index."""
    if source_kind == target_kind == "vba":
        return "call"
    if source_kind == "vba":
        return "vba-write"
    if target_kind == "vba":
        return "vba-read"
    return "dep"


def static_edge_record(source, target, source_kind, target_kind):
    """Preserve static VBA uncertainty when reading a legacy edge index."""
    kind = static_edge_kind(source_kind, target_kind)
    relation = {"source": source, "target": target, "kind": kind}
    if kind != "dep":
        relation.update(scope="static_vba_heuristic", approx=True)
    return relation
