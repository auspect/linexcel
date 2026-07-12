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
