"""High-level API, usable as a library (marimo, Jupyter, scripts).

Minimal example, without backend or AI key:

    from linexcel import analyze
    result = analyze("my_workbook.xlsx")
    result                      # interactive graph in marimo
    result.save_html("out.html")
    print(result.stats)

AI documentation is optional:

    result.document(api_key="...")   # or via GOOGLE_API_KEY env var
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
    """Normalize path / bytes / file object into (bytes, filename)."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source), filename or "workbook.xlsx"
    if isinstance(source, (str, Path)):
        path = Path(source)
        return path.read_bytes(), filename or path.name
    if hasattr(source, "read"):
        data = source.read()
        if not isinstance(data, (bytes, bytearray)):
            raise TypeError("Stream must be opened in binary mode ('rb').")
        name = filename or getattr(source, "name", None) or "workbook.xlsx"
        return bytes(data), Path(str(name)).name
    raise TypeError("source must be a path, bytes, or a binary file object.")


def analyze(source: Source, filename: str | None = None) -> LineageResult:
    """Analyze an Excel workbook and return a :class:`LineageResult`.

    Parameters
    ----------
    source : str | Path | bytes | binary file
        Path to the file, raw content, or file object opened in ``rb``.
    filename : str, optional
        Logical name (used for labels and VBA detection).
    """
    data, name = _read_source(source, filename)
    payload = analyze_workbook(data, filename=name)
    return LineageResult(
        graph=payload["graph"],
        engine=payload["engine"],
        analysis_id=payload["analysisId"],
        source_data=data,
        filename=name,
    )


class LineageResult:
    """Analysis result: deterministic graph + computation engine + renderers.

    The object is directly displayable in a notebook (``_repr_html_``) and
    exposes the JSON graph, convenience accessors, standalone HTML export,
    and optional AI documentation.
    """

    def __init__(
        self,
        graph: dict[str, Any],
        engine: Any,
        analysis_id: str | None = None,
        source_data: bytes | None = None,
        filename: str | None = None,
    ):
        self.graph = graph
        self.engine = engine
        self.analysis_id = analysis_id or uuid.uuid4().hex[:16]
        self._by_id = {n["id"]: n for n in graph.get("nodes", [])}
        self._source_data = source_data
        self._source_filename = filename or graph.get("meta", {}).get(
            "filename", "workbook.xlsx"
        )
        self._workbook_context: dict[str, Any] | None = None

    # -- convenience accessors --------------------------------------------
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

    @property
    def workbook_context(self) -> dict[str, Any]:
        """Bounded sheet previews, comments, and layout markers.

        Context is extracted with ``openpyxl`` only; Excel or LibreOffice is
        not launched. It deliberately preserves first rows and columns rather
        than assuming a tabular header convention.
        """
        if self._workbook_context is None:
            from linexcel.insights import extract_workbook_context

            self._workbook_context = extract_workbook_context(
                self._source_bytes(), self._source_filename
            )
        return self._workbook_context

    def node(self, node_id: str) -> dict[str, Any] | None:
        """Return the node with the given id (or ``None``)"""
        return self._by_id.get(node_id)

    def find(self, text: str) -> list[dict[str, Any]]:
        """
        Nodes whose label or formula contains ``text`` (case-insensitive)
        """
        q = text.lower()
        return [
            n
            for n in self.nodes
            if q in (n.get("label", "").lower())
            or q in (n.get("formula", "") or "").lower()
        ]

    def precedents(self, node_id: str) -> list[dict[str, Any]]:
        """Nodes that feed into ``node_id``"""
        return [
            self._by_id[e["source"]]
            for e in self.edges
            if e["target"] == node_id and e["source"] in self._by_id
        ]

    def dependents(self, node_id: str) -> list[dict[str, Any]]:
        """Nodes fed by ``node_id``"""
        return [
            self._by_id[e["target"]]
            for e in self.edges
            if e["source"] == node_id and e["target"] in self._by_id
        ]

    # -- serialization -----------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        return self.graph

    def to_json(self, *, indent: int | None = None) -> str:
        return json.dumps(
            self.graph, ensure_ascii=False, indent=indent, default=str
        )

    def save_json(self, path: str | Path) -> Path:
        path = Path(path)
        path.write_text(self.to_json(indent=1), encoding="utf-8")
        return path

    def save_screenshots(
        self,
        output_dir: str | Path,
        *,
        dpi: int = 144,
        timeout: int = 60,
    ) -> list[Path]:
        """Render workbook pages to PNG using LibreOffice headless on Linux.

        The optional renderer requires ``libreoffice`` (or ``soffice``) and
        ``pdftoppm`` from Poppler. Use :attr:`workbook_context` when only the
        non-rendered context is needed.
        """
        from linexcel.insights import render_workbook_screenshots

        return render_workbook_screenshots(
            self._source_bytes(),
            self._source_filename,
            output_dir,
            dpi=dpi,
            timeout=timeout,
        )

    # -- AI documentation (optional) --------------------------------------
    def document(
        self,
        node_ids: list[str] | None = None,
        *,
        api_key: str | None = None,
        model: str | None = None,
        language: str = "en",
    ) -> dict[str, str]:
        """Document nodes via Gemini from the deterministic lineage.

        Without ``node_ids``, documents all calculation nodes
        (cells, groups, VBA).
        Requires ``google-genai`` and a key
        (``api_key`` or ``GOOGLE_API_KEY``).

        ``language`` selects the system prompt ("en" or "fr").
        """
        from linexcel.aidoc import document_nodes

        if node_ids is None:
            node_ids = [
                n["id"]
                for n in self.nodes
                if n.get("kind") in ("cell", "group", "vba")
            ]
        return document_nodes(
            self.graph,
            node_ids,
            model=model,
            api_key=api_key,
            language=language,
        )

    # -- visualization -----------------------------------------------------
    def to_html(
        self,
        *,
        title: str | None = None,
        full_document: bool = True,
        docs: dict[str, str] | None = None,
    ) -> str:
        """Standalone HTML document (Cytoscape) — openable in a browser.

        If ``docs`` is provided (from :meth:`document`), AI documentation
        for each node is embedded in the detail panel.
        """
        graph = self.graph
        if docs:
            graph = {
                **graph,
                "nodes": [
                    {**n, "doc": docs.get(n["id"], "")} for n in graph["nodes"]
                ],
            }
        return render_html(
            graph, title=title or self._title(), full_document=full_document
        )

    def save_html(
        self,
        path: str | Path,
        *,
        title: str | None = None,
        docs: dict[str, str] | None = None,
    ) -> Path:
        path = Path(path)
        path.write_text(self.to_html(title=title, docs=docs), encoding="utf-8")
        return path

    def _title(self) -> str:
        return self.graph.get("meta", {}).get("filename", "Lineage Excel")

    def _source_bytes(self) -> bytes:
        if self._source_data is None:
            raise RuntimeError(
                "Workbook bytes are unavailable. Create the result with analyze()."
            )
        return self._source_data

    def _repr_html_(self) -> str:
        """Inline rendering for marimo / Jupyter (isolated iframe)."""
        return wrap_iframe(self.to_html(), height=640)

    def __repr__(self) -> str:
        s = self.stats
        return (
            f"<LineageResult {self._title()!r}: "
            f"{s['totalFormulas']} formulas, {s['totalNodes']} nodes, "
            f"{s['totalEdges']} edges, {s['vbaProcs']} VBA procs>"
        )
