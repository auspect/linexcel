"""API haut-niveau, utilisable comme bibliothèque (marimo, Jupyter, scripts).

Exemple minimal, sans backend ni clé IA :

    from linexcel import analyze
    result = analyze("mon_classeur.xlsx")
    result                      # s'affiche en graphe interactif dans marimo
    result.save_html("out.html")
    print(result.stats)

La documentation IA reste optionnelle :

    result.document(api_key="…")   # ou via la variable GOOGLE_API_KEY
"""

from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any, BinaryIO

from linexcel.analyzer import analyze_workbook
from linexcel.viewer import render_html, wrap_iframe

Source = str | Path | bytes | bytearray | BinaryIO


def _read_source(source: Source, filename: str | None) -> tuple[bytes, str]:
    """Normalise chemin / bytes / objet fichier en (octets, nom de fichier)."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source), filename or "classeur.xlsx"
    if isinstance(source, (str, Path)):
        path = Path(source)
        return path.read_bytes(), filename or path.name
    if hasattr(source, "read"):
        data = source.read()
        if not isinstance(data, (bytes, bytearray)):
            raise TypeError("Le flux doit être ouvert en mode binaire ('rb').")
        name = filename or getattr(source, "name", None) or "classeur.xlsx"
        return bytes(data), Path(str(name)).name
    raise TypeError(
        "source doit être un chemin, des octets ou un fichier binaire ouvert."
    )


def analyze(source: Source, filename: str | None = None) -> LineageResult:
    """Analyse un classeur Excel et retourne un :class:`LineageResult`.

    Paramètres
    ----------
    source : str | Path | bytes | fichier binaire
        Chemin vers le fichier, contenu brut, ou objet ouvert en ``rb``.
    filename : str, optionnel
        Nom logique (sert aux libellés et à la détection VBA).
    """
    data, name = _read_source(source, filename)
    payload = analyze_workbook(data, filename=name)
    return LineageResult(
        graph=payload["graph"],
        engine=payload["engine"],
        analysis_id=payload["analysisId"],
    )


class LineageResult:
    """Résultat d'analyse : graphe déterministe + moteur de calcul + rendus.

    L'objet est directement affichable dans un notebook (``_repr_html_``) et
    expose le graphe JSON, des accès pratiques, l'export HTML autonome et la
    documentation IA optionnelle.
    """

    def __init__(
        self, graph: dict[str, Any], engine: Any, analysis_id: str | None = None
    ):
        self.graph = graph
        self.engine = engine
        self.analysis_id = analysis_id or uuid.uuid4().hex[:16]
        self._by_id = {n["id"]: n for n in graph.get("nodes", [])}

    # -- accès pratiques ---------------------------------------------------
    @property
    def nodes(self) -> list[dict[str, Any]]:
        return self.graph["nodes"]

    @property
    def edges(self) -> list[dict[str, Any]]:
        return self.graph["edges"]

    @property
    def sheets(self) -> list[str]:
        return self.graph.get("sheets", [])

    @property
    def stats(self) -> dict[str, Any]:
        return self.graph["meta"]["stats"]

    @property
    def warnings(self) -> list[str]:
        return self.graph["meta"]["warnings"]

    def node(self, node_id: str) -> dict[str, Any] | None:
        """Renvoie le nœud d'identifiant donné (ou ``None``)."""
        return self._by_id.get(node_id)

    def find(self, text: str) -> list[dict[str, Any]]:
        """Nœuds dont le libellé ou la formule contient ``text`` (insensible casse)."""
        q = text.lower()
        return [
            n
            for n in self.nodes
            if q in (n.get("label", "").lower())
            or q in (n.get("formula", "") or "").lower()
        ]

    def precedents(self, node_id: str) -> list[dict[str, Any]]:
        """Nœuds qui alimentent ``node_id``."""
        return [
            self._by_id[e["source"]]
            for e in self.edges
            if e["target"] == node_id and e["source"] in self._by_id
        ]

    def dependents(self, node_id: str) -> list[dict[str, Any]]:
        """Nœuds alimentés par ``node_id``."""
        return [
            self._by_id[e["target"]]
            for e in self.edges
            if e["source"] == node_id and e["target"] in self._by_id
        ]

    # -- sérialisation -----------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        return self.graph

    def to_json(self, *, indent: int | None = None) -> str:
        return json.dumps(self.graph, ensure_ascii=False, indent=indent, default=str)

    def save_json(self, path: str | Path) -> Path:
        path = Path(path)
        path.write_text(self.to_json(indent=1), encoding="utf-8")
        return path

    # -- documentation IA (optionnelle) -----------------------------------
    def document(
        self,
        node_ids: list[str] | None = None,
        *,
        api_key: str | None = None,
        model: str | None = None,
    ) -> dict[str, str]:
        """Documente les nœuds via Gemini à partir du lignage déterministe.

        Sans ``node_ids``, documente tous les calculs (cellules, groupes, VBA).
        Nécessite ``google-genai`` et une clé (``api_key`` ou ``GOOGLE_API_KEY``).
        """
        from linexcel.aidoc import document_nodes

        if node_ids is None:
            node_ids = [
                n["id"] for n in self.nodes if n.get("kind") in ("cell", "group", "vba")
            ]
        return document_nodes(self.graph, node_ids, model=model, api_key=api_key)

    # -- visualisation -----------------------------------------------------
    def to_html(
        self, *, title: str | None = None, full_document: bool = True,
        docs: dict[str, str] | None = None,
    ) -> str:
        """Document HTML autonome (Cytoscape) — ouvrable dans un navigateur.

        Si ``docs`` est fourni (issu de :meth:`document`), la documentation IA
        de chaque nœud est embarquée dans le panneau de détail.
        """
        graph = self.graph
        if docs:
            graph = {**graph, "nodes": [
                {**n, "doc": docs.get(n["id"], "")} for n in graph["nodes"]
            ]}
        return render_html(
            graph, title=title or self._title(), full_document=full_document
        )

    def save_html(
        self, path: str | Path, *, title: str | None = None,
        docs: dict[str, str] | None = None,
    ) -> Path:
        path = Path(path)
        path.write_text(self.to_html(title=title, docs=docs), encoding="utf-8")
        return path

    def _title(self) -> str:
        return self.graph.get("meta", {}).get("filename", "Lineage Excel")

    def _repr_html_(self) -> str:
        """Rendu inline pour marimo / Jupyter (iframe isolée)."""
        return wrap_iframe(self.to_html(), height=640)

    def __repr__(self) -> str:
        s = self.stats
        return (
            f"<LineageResult {self._title()!r}: "
            f"{s['totalFormulas']} formules, {s['totalNodes']} nœuds, "
            f"{s['totalEdges']} liens, {s['vbaProcs']} proc. VBA>"
        )
