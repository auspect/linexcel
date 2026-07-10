"""Documentation automatique des calculs par IA (google-genai / Gemini).

Le modèle ne « devine » rien : chaque nœud lui est présenté avec son dossier
déterministe issu du graphe (formule exacte, décomposition évaluée étape par
étape, précédents et leurs valeurs, dépendants, étendue du groupe étiré,
liens VBA). La consigne impose de ne citer que ces faits, ce qui rend la
documentation « prouvable » : chaque affirmation renvoie à la formule ou à
une valeur du classeur.
"""

from __future__ import annotations

import json
import os
from typing import Any

DEFAULT_MODEL = "gemini-2.5-flash"
MAX_DOSSIER_CHARS = 6_000

_SYSTEM = """Tu documentes des calculs Excel pour un lecteur métier francophone.
Pour le nœud fourni, rédige une fiche courte en Markdown :
1. **Rôle** — une phrase sur ce que calcule la formule ;
2. **Comment** — la logique, étape par étape, en t'appuyant STRICTEMENT sur la
   décomposition fournie (cite les sous-expressions et leurs valeurs évaluées) ;
3. **Sources** — d'où viennent les données (précédents, plages, noms, VBA) ;
4. **Preuve** — la formule exacte et, si disponible, la valeur calculée.
Règles absolues : n'invente aucune donnée ; n'affirme rien qui ne soit pas dans
le dossier ; si une information manque, écris « non déterminé par le lignage ».
Réponds UNIQUEMENT avec la fiche Markdown, aucun JSON, aucun délimiteur."""


class AiDocError(RuntimeError):
    pass


def _client(api_key: str | None = None):
    try:
        from google import genai
    except ImportError as exc:  # pragma: no cover
        raise AiDocError(
            "Le paquet google-genai n'est pas installé "
            "(pip install 'backend[ai]' ou pip install google-genai)"
        ) from exc
    api_key = api_key or os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise AiDocError(
            "Aucune clé Gemini fournie : passer api_key=... ou définir "
            "GOOGLE_API_KEY dans l'environnement"
        )
    return genai.Client(api_key=api_key)


def build_dossier(graph: dict[str, Any], node_id: str) -> dict[str, Any] | None:
    """Dossier déterministe d'un nœud : tout ce que l'IA a le droit d'utiliser."""
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
        "adresse": node.get("addr"),
        "formule": node.get("formula"),
        "forme_r1c1": node.get("r1c1"),
        "cellules_du_groupe": node.get("count"),
        "etendue": node.get("bbox"),
        "valeur_calculee": node.get("value"),
        "exemples_valeurs": node.get("samples"),
        "decomposition": _compact_steps(node.get("steps")),
        "precedents": precedents[:30],
        "dependants": dependents[:30],
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
        "lien": edge.get("kind"),
        "valeur": other.get("value"),
        "formule": other.get("formula"),
    }


def _compact_steps(step: dict | None) -> dict | None:
    if step is None:
        return None
    out = {
        "expression": step.get("expr"),
        "operation": step.get("label"),
        "valeur": step.get("value") if step.get("evaluated") else "non évaluée",
    }
    if step.get("inputs"):
        out["entrees"] = step["inputs"]
    children = [_compact_steps(c) for c in step.get("children", [])]
    if children:
        out["sous_etapes"] = children
    return out


def document_nodes(
    graph: dict[str, Any],
    node_ids: list[str],
    model: str | None = None,
    api_key: str | None = None,
) -> dict[str, str]:
    """Documente les nœuds demandés, par lots, et retourne {node_id: markdown}."""
    client = _client(api_key)
    model = model or os.getenv("GEMINI_MODEL", DEFAULT_MODEL)
    docs: dict[str, str] = {}
    dossiers = []
    for nid in node_ids:
        d = build_dossier(graph, nid)
        if d is not None:
            blob = json.dumps(d, ensure_ascii=False, default=str)
            if len(blob) > MAX_DOSSIER_CHARS:
                d["decomposition"] = "tronquée (formule très longue)"
                blob = json.dumps(d, ensure_ascii=False, default=str)
            dossiers.append((nid, blob))

    for nid, blob in dossiers:
        prompt = (
            _SYSTEM
            + "\n\nDossier de lignage (déterministe, extrait du classeur) :\n"
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
            raise AiDocError(f"Appel Gemini en échec : {exc}") from exc
        if text:
            docs[nid] = text
    return docs
