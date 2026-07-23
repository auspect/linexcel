"""AI-generated documentation for Excel calculations.

Multi-provider support via a thin abstraction:
- Google Gemini (google-genai) — default, backward-compatible
- OpenAI-compatible API (any endpoint: OpenAI, Ollama, vLLM, LM Studio, etc.)
- Callable protocol for custom/local models

The model doesn't guess: each node is presented with its deterministic dossier
from the graph (exact formula, step-by-step evaluation, precedents and their
values, dependents, stretched group extent, VBA links). The system prompt
enforces citing only these facts, making the documentation "provable": every
claim traces back to a formula or a workbook value.
"""

from __future__ import annotations

import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Protocol, runtime_checkable

DEFAULT_MODEL = "gemini-3.1-flash-lite"
MAX_DOSSIER_CHARS = 6_000
MAX_WORKBOOK_DOSSIER_CHARS = 12_000

_LANGUAGES = ("en", "fr")

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


# ──────────────────────────────────────────────
# Provider abstraction
# ──────────────────────────────────────────────


@runtime_checkable
class LLMProvider(Protocol):
    """Minimal protocol: system + user prompt → text response."""

    def generate(self, system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
        ...


def _resolve_provider(
    *,
    provider: LLMProvider | None = None,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
) -> tuple[LLMProvider, str]:
    """Resolve which provider to use and the model name.

    Priority:
    1. Explicit `provider` (custom callable or LLMProvider instance)
    2. `base_url` set → OpenAI-compatible client
    3. Google API key available → Gemini client (backward-compatible default)
    """
    if provider is not None:
        return provider, model or DEFAULT_MODEL

    # OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, OpenAI, etc.)
    base = base_url or os.getenv("LINEXCEL_AI_BASE_URL") or os.getenv("OPENAI_BASE_URL")
    if base:
        resolved_model = model or os.getenv("LINEXCEL_AI_MODEL") or os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
        return _OpenAICompatProvider(base_url=base, api_key=api_key, model=resolved_model), resolved_model

    # Google Gemini (default, backward-compatible)
    resolved_model = model or os.getenv("GEMINI_MODEL", DEFAULT_MODEL)
    return _GeminiProvider(api_key=api_key, model=resolved_model), resolved_model


class _GeminiProvider:
    """Google Gemini via google-genai."""

    def __init__(self, *, api_key: str | None = None, model: str = DEFAULT_MODEL):
        try:
            from google import genai
        except ImportError as exc:  # pragma: no cover
            raise AiDocError(
                "google-genai is not installed "
                "(pip install 'linexcel[ai]' or pip install google-genai)"
            ) from exc
        self._genai = genai
        self._api_key = (
            api_key
            or os.getenv("GOOGLE_API_KEY")
            or os.getenv("GEMINI_API_KEY")
        )
        if not self._api_key:
            raise AiDocError(
                "No Gemini API key provided: pass api_key=... or set "
                "GOOGLE_API_KEY in the environment"
            )
        self._client = genai.Client(api_key=self._api_key)
        self._model = model

    def generate(self, system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
        try:
            response = self._client.models.generate_content(
                model=self._model,
                contents=system_prompt + "\n\n" + user_prompt,
                config={"temperature": temperature},
            )
            return (response.text or "").strip()
        except Exception as exc:
            raise AiDocError(f"Gemini API call failed: {exc}") from exc


class _OpenAICompatProvider:
    """OpenAI-compatible API client (works with Ollama, vLLM, LM Studio, etc.)."""

    def __init__(self, *, base_url: str, api_key: str | None = None, model: str = "gpt-4o-mini"):
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover
            raise AiDocError(
                "openai is not installed (pip install openai)"
            ) from exc
        self._client = OpenAI(
            api_key=api_key or os.getenv("OPENAI_API_KEY") or "ollama",  # ollama doesn't need a real key
            base_url=base_url,
        )
        self._model = model

    def generate(self, system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=temperature,
            )
            return (response.choices[0].message.content or "").strip()
        except Exception as exc:
            raise AiDocError(f"OpenAI-compatible API call failed: {exc}") from exc


# ──────────────────────────────────────────────
# Dossier builders (unchanged)
# ──────────────────────────────────────────────


def build_dossier(graph: dict[str, Any], node_id: str) -> dict[str, Any] | None:
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


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────


def document_workbook(
    graph: dict[str, Any],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: LLMProvider | None = None,
    language: str = "en",
) -> str:
    """Generate a Markdown overview grounded in the workbook dossier.

    Provider resolution (first match wins):
    1. `provider` — custom LLMProvider instance or callable
    2. `base_url` or `LINEXCEL_AI_BASE_URL` — OpenAI-compatible endpoint
    3. Google Gemini (default, backward-compatible)
    """
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
    llm, resolved_model = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
    system = _WORKBOOK_SYSTEM[language]
    user = (
        "Workbook dossier (deterministic, extracted from workbook):\n" + blob
    )
    try:
        return llm.generate(system, user, temperature=0.2)
    except AiDocError:
        raise
    except Exception as exc:
        raise AiDocError(f"AI documentation failed: {exc}") from exc


def document_nodes(
    graph: dict[str, Any],
    node_ids: list[str],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: LLMProvider | None = None,
    language: str = "en",
) -> dict[str, str]:
    """Document the requested nodes, returns {node_id: markdown}.

    Provider resolution is the same as :func:`document_workbook`.
    """
    if language not in _LANGUAGES:
        raise ValueError(
            f"Unsupported language: {language!r}. Use one of {_LANGUAGES}"
        )
    llm, resolved_model = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
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

    def _doc_one(nid_blob: tuple[str, str]) -> tuple[str, str]:
        nid, blob = nid_blob
        user = "Lineage dossier (deterministic, extracted from workbook):\n" + blob
        text = llm.generate(system, user, temperature=0.2)
        return nid, text or "(AI returned empty response)"

    # ponytail: 4 workers, bump if API rate limits allow
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(_doc_one, d): d[0] for d in dossiers}
        for fut in as_completed(futures):
            node_id = futures[fut]
            try:
                nid, text = fut.result()
            except Exception as exc:
                raise AiDocError(
                    f"AI documentation failed for node {node_id}: {exc}"
                ) from exc
            docs[nid] = text
    return docs
