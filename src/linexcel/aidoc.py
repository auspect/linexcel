"""AI-generated documentation for Excel calculations (google-genai / Gemini).

The model doesn't guess: each node is presented with its deterministic dossier
from the graph (exact formula, step-by-step evaluation, precedents and their
values, dependents, stretched group extent, VBA links). The system prompt
enforces citing only these facts, making the documentation "provable": every
claim traces back to a formula or a workbook value.
"""

from __future__ import annotations

import json
import os
from typing import Any

DEFAULT_MODEL = "gemini-3.1-flash-lite"
MAX_DOSSIER_CHARS = 6_000
MAX_WORKBOOK_DOSSIER_CHARS = 12_000

_SYSTEM = {
    "en": """
You document Excel calculations for a business reader.
For the provided node, write a short Markdown card:
1. **Role** — one sentence on what the formula computes;
2. **How** — the logic, step by step, relying STRICTLY on the provided
   decomposition (cite sub-expressions and their evaluated values);
3. **Sources** — where the data comes from (precedents, ranges, names, VBA);
4. **Proof** — the exact formula and, if available, the computed value.
Absolute rules: do not invent data; do not assert anything not in the
dossier; if information is missing, write "not determined by lineage".
Respond ONLY with the Markdown card, no JSON, no delimiters.
""".strip(),
    "fr": """
Tu documentes des calculs Excel pour un lecteur métier francophone.
Pour le nœud fourni, rédige une fiche courte en Markdown :
1. **Rôle** — une phrase sur ce que calcule la formule ;
2. **Comment** — la logique, étape par étape, en t'appuyant STRICTEMENT sur la
   décomposition fournie (cite les sous-expressions et leurs valeurs évaluées) ;
3. **Sources** — d'où viennent les données (précédents, plages, noms, VBA) ;
4. **Preuve** — la formule exacte et, si disponible, la valeur calculée.
Règles absolues : n'invente aucune donnée ; n'affirme rien qui ne soit pas dans
le dossier ; si une information manque, écris « non déterminé par le lignage ».
Réponds UNIQUEMENT avec la fiche Markdown, aucun JSON, aucun délimiteur.
""".strip(),
}

_LANGUAGES = ("en", "fr")

_WORKBOOK_SYSTEM = {
    "en": """
You document an Excel workbook for a business reader.
Write a concise Markdown overview with these sections:
1. **Purpose** — the workbook's apparent role, only when supported by the dossier;
2. **Structure** — its sheets and how calculations are distributed;
3. **Calculation flow** — important formula patterns, defined names, and links;
4. **Automation and caveats** — VBA, external references, warnings, and analysis limits;
5. **Questions to validate** — up to five concrete items that cannot be determined.
Use only facts in the deterministic dossier. Do not infer a business purpose from
sheet names alone. State "not determined by lineage" for missing information.
Respond ONLY with the Markdown overview, no JSON or delimiters.
""".strip(),
    "fr": """
Tu documentes un classeur Excel pour un lecteur métier.
Rédige une synthèse concise en Markdown avec les sections suivantes :
1. **Rôle** — la fonction apparente du classeur, uniquement si le dossier le confirme ;
2. **Structure** — ses feuilles et la répartition des calculs ;
3. **Flux de calcul** — les principaux motifs de formules, noms définis et liens ;
4. **Automatisation et limites** — VBA, références externes,
   avertissements et limites d'analyse ;
5. **Questions à valider** — au plus cinq points concrets indéterminables.
Utilise uniquement les faits présents dans le dossier déterministe. N'infère pas
un rôle métier à partir des seuls noms de feuilles. Écris « non déterminé par le
lignage » lorsqu'une information manque. Réponds UNIQUEMENT avec la synthèse
Markdown, sans JSON ni délimiteur.
""".strip(),
}


class AiDocError(RuntimeError):
    pass


def _client(api_key: str | None = None):
    try:
        from google import genai
    except ImportError as exc:  # pragma: no cover
        raise AiDocError(
            "google-genai is not installed "
            "(pip install 'linexcel[ai]' or pip install google-genai)"
        ) from exc
    api_key = (
        api_key or os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
    )
    if not api_key:
        raise AiDocError(
            "No Gemini API key provided: pass api_key=... or set "
            "GOOGLE_API_KEY in the environment"
        )
    return genai.Client(api_key=api_key)


def build_dossier(
    graph: dict[str, Any], node_id: str
) -> dict[str, Any] | None:
    """
    Deterministic dossier for a node: everything the AI is allowed to use.
    """
    nodes = {n["id"]: n for n in graph["nodes"]}
    node = nodes.get(node_id)
    if node is None:
        return None
    precedents, dependents = [], []
    for e in graph["edges"]:
        if e["target"] == node_id:
            src = nodes.get(e["source"], {})
            precedents.append(_neighbor(src, e))
        elif e["source"] == node_id:
            dst = nodes.get(e["target"], {})
            dependents.append(_neighbor(dst, e))
    dossier = {
        "node_id": node_id,
        "kind": node.get("kind"),
        "sheet": node.get("sheet"),
        "address": node.get("addr"),
        "formula": node.get("formula"),
        "r1c1_form": node.get("r1c1"),
        "group_cells": node.get("count"),
        "extent": node.get("bbox"),
        "computed_value": node.get("value"),
        "value_samples": node.get("samples"),
        "decomposition": _compact_steps(node.get("steps")),
        "precedents": precedents[:30],
        "dependents": dependents[:30],
    }
    if node.get("kind") == "vba":
        dossier["vba"] = {
            "module": node.get("module"),
            "procedure": node.get("proc"),
            "type": node.get("procKind"),
            "code": (node.get("code") or "")[:2500],
        }
    return dossier


def _neighbor(other: dict, edge: dict) -> dict:
    return {
        "id": other.get("id"),
        "kind": other.get("kind"),
        "label": other.get("label"),
        "edge_kind": edge.get("kind"),
        "value": other.get("value"),
        "formula": other.get("formula"),
    }


def _compact_steps(step: dict | None) -> dict | None:
    if step is None:
        return None
    out = {
        "expression": step.get("expr"),
        "operation": step.get("label"),
        "value": step.get("value")
        if step.get("evaluated")
        else "not evaluated",
    }
    if step.get("inputs"):
        out["inputs"] = step["inputs"]
    children = [_compact_steps(c) for c in step.get("children", [])]
    if children:
        out["sub_steps"] = children
    return out


def build_workbook_dossier(graph: dict[str, Any]) -> dict[str, Any]:
    """Return a compact, deterministic dossier for a whole-workbook overview."""
    nodes = graph.get("nodes", [])
    meta = graph.get("meta", {})
    stats = meta.get("stats", {})
    sheet_stats = stats.get("sheets", [])
    nodes_by_sheet: dict[str, dict[str, int]] = {}
    for node in nodes:
        sheet = node.get("sheet")
        if not sheet:
            continue
        kinds = nodes_by_sheet.setdefault(sheet, {})
        kind = node.get("kind", "unknown")
        kinds[kind] = kinds.get(kind, 0) + 1

    sheets = [
        {
            "name": sheet.get("name"),
            "dimensions": {"rows": sheet.get("rows"), "columns": sheet.get("cols")},
            "formula_cells": sheet.get("formulaCells", 0),
            "lineage_nodes": nodes_by_sheet.get(sheet.get("name"), {}),
        }
        for sheet in sheet_stats
    ]
    formula_patterns = sorted(
        (
            {
                "sheet": node.get("sheet"),
                "address": node.get("addr"),
                "formula": node.get("formula"),
                "cells": node.get("count", 1),
                "extent": node.get("bbox"),
            }
            for node in nodes
            if node.get("kind") in {"cell", "group"}
        ),
        key=lambda item: item["cells"],
        reverse=True,
    )[:20]
    defined_names = [
        {"name": node.get("label"), "targets": node.get("targets", [])}
        for node in nodes
        if node.get("kind") == "name"
    ]
    vba = [
        {
            "module": node.get("module"),
            "procedure": node.get("proc"),
            "type": node.get("procKind"),
        }
        for node in nodes
        if node.get("kind") == "vba"
    ]
    opaque_references = [
        node.get("label") for node in nodes if node.get("kind") == "opaque"
    ]
    return {
        "filename": meta.get("filename"),
        "analysis": {
            "formula_cells": stats.get("totalFormulas", 0),
            "lineage_nodes": stats.get("totalNodes", 0),
            "lineage_edges": stats.get("totalEdges", 0),
            "grouped_patterns": stats.get("groupedPatterns", 0),
        },
        "sheets": sheets,
        "formula_patterns": formula_patterns,
        "defined_names": defined_names,
        "vba_procedures": vba,
        "external_or_unresolved_references": opaque_references,
        "warnings": meta.get("warnings", []),
    }


def document_workbook(
    graph: dict[str, Any],
    *,
    model: str | None = None,
    api_key: str | None = None,
    language: str = "en",
) -> str:
    """Generate a Markdown overview grounded in the workbook dossier."""
    if language not in _LANGUAGES:
        raise ValueError(
            f"Unsupported language: {language!r}. Use one of {_LANGUAGES}"
        )
    dossier = build_workbook_dossier(graph)
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) > MAX_WORKBOOK_DOSSIER_CHARS:
        dossier["formula_patterns"] = dossier["formula_patterns"][:5]
        dossier["vba_procedures"] = dossier["vba_procedures"][:10]
        blob = json.dumps(dossier, ensure_ascii=False, default=str)
    client = _client(api_key)
    try:
        response = client.models.generate_content(
            model=model or os.getenv("GEMINI_MODEL", DEFAULT_MODEL),
            contents=(
                _WORKBOOK_SYSTEM[language]
                + "\n\nWorkbook dossier (deterministic, extracted from workbook):\n"
                + blob
            ),
            config={"temperature": 0.2},
        )
        return (response.text or "").strip()
    except AiDocError:
        raise
    except Exception as exc:
        raise AiDocError(f"Gemini API call failed: {exc}") from exc


def document_nodes(
    graph: dict[str, Any],
    node_ids: list[str],
    *,
    model: str | None = None,
    api_key: str | None = None,
    language: str = "en",
) -> dict[str, str]:
    """Document the requested nodes in batches, returns {node_id: markdown}.

    ``language`` selects the system prompt ("en" or "fr").
    """
    if language not in _LANGUAGES:
        raise ValueError(
            f"Unsupported language: {language!r}. Use one of {_LANGUAGES}"
        )
    client = _client(api_key)
    model = model or os.getenv("GEMINI_MODEL", DEFAULT_MODEL)
    system = _SYSTEM[language]
    docs: dict[str, str] = {}
    dossiers = []
    for nid in node_ids:
        d = build_dossier(graph, nid)
        if d is not None:
            blob = json.dumps(d, ensure_ascii=False, default=str)
            if len(blob) > MAX_DOSSIER_CHARS:
                d["decomposition"] = "truncated (very long formula)"
                blob = json.dumps(d, ensure_ascii=False, default=str)
            dossiers.append((nid, blob))

    for nid, blob in dossiers:
        prompt = (
            system
            + "\n\nLineage dossier (deterministic, extracted from workbook):\n"
            + blob
        )
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config={"temperature": 0.2},
            )
            text = (response.text or "").strip()
        except AiDocError:
            raise
        except Exception as exc:
            raise AiDocError(f"Gemini API call failed: {exc}") from exc
        if text:
            docs[nid] = text
    return docs
