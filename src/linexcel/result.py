"""High-level API, usable as a library (marimo, Jupyter, scripts).

Minimal example, without backend or AI key:

    from linexcel import analyze
    result = analyze("my_workbook.xlsx")
    result                      # interactive graph in marimo
    result.save_html("out.html")
    print(result.stats)

AI documentation is optional — pick a provider explicitly (no default):

    result.document(base_url="http://localhost:11434/v1", model="qwen3.8")
"""

from __future__ import annotations

import json
import uuid
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO, cast

from linexcel.analyzer import analyze_workbook
from linexcel.viewer import render_html, wrap_iframe

if TYPE_CHECKING:  # aidoc stays lazily imported at runtime
    from linexcel.aidoc import ProviderLike, TokenUsage

Source = str | Path | bytes | bytearray | BinaryIO
# Covariant containers: callers commonly hold a list[Path] / dict[str, list[Path]].
Screenshots = Sequence[str | Path] | Mapping[str, Sequence[str | Path]]


def _first_shot(shots: str | Path | Sequence[str | Path]) -> str | Path | None:
    """The image standing for one sheet, whether it came alone or in a list.

    A lone ``str`` is a path, not a sequence of characters: indexing it would
    hand the model the letter ``C`` of ``C:\\shots\\Sales.png``.
    """
    if isinstance(shots, (str, Path)):
        return shots
    listed = list(shots)
    return listed[0] if listed else None


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


def analyze(
    source: Source,
    filename: str | None = None,
    *,
    verbose: bool = False,
    refs_dir: str | Path | None = None,
) -> LineageResult:
    """Analyze an Excel workbook and return a :class:`LineageResult`.

    Parameters
    ----------
    source : str | Path | bytes | binary file
        Path to the file, raw content, or file object opened in ``rb``.
    filename : str, optional
        Logical name (used for labels and VBA detection).
    verbose : bool, optional
        Print per-phase timing to stderr.
    refs_dir : str | Path, optional
        Folder holding the workbooks this one links to, and the add-ins whose
        VBA it calls. Without it a cell reading ``'[Budget.xlsx]Annual'!B4`` is
        named but left unresolved; with it the referenced file is read and the
        reference evaluates to the value it stands for. The report states, per
        workbook, whether it was read from the folder, taken from the cache
        Excel left in the file, or not read at all.
    """
    data, name = _read_source(source, filename)
    try:
        payload = analyze_workbook(
            data, filename=name, verbose=verbose, refs_dir=refs_dir
        )
    except Exception as exc:
        # Frontière publique : transformer l'erreur brute (BadZipFile, Rust)
        # en message clair sur le vrai problème.
        if not data[:4] == b"PK\x03\x04":
            raise ValueError(
                f"{name!r} is not an Excel file (xlsx/xlsm). "
                "Legacy .xls is not supported — re-save it as .xlsx first."
            ) from exc
        raise ValueError(f"Could not analyze {name!r}: {exc}") from exc
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
        self._token_usage: TokenUsage | None = None

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

    @property
    def token_usage(self) -> TokenUsage:
        """Tokens consumed by every :meth:`document` / :meth:`document_workbook`
        call made on this result.

        Counts come from the provider when it reports them (an
        OpenAI-compatible endpoint fills in a ``usage`` block); otherwise they
        are approximated and :attr:`TokenUsage.estimated` is set. Zero until an
        AI call is made.

            >>> result.document(provider=my_llm)          # doctest: +SKIP
            >>> print(result.token_usage)                 # doctest: +SKIP
            ~12,480 tokens (~11,120 in + ~1,360 out) over 4 request(s)
        """
        if self._token_usage is None:
            from linexcel.aidoc import TokenUsage as _TokenUsage

            self._token_usage = _TokenUsage()
        return self._token_usage

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
        return json.dumps(self.graph, ensure_ascii=False, indent=indent, default=str)

    def save_json(self, path: str | Path) -> Path:
        path = Path(path)
        path.write_text(self.to_json(indent=1), encoding="utf-8")
        return path

    def save_screenshots(
        self,
        output_dir: str | Path,
        *,
        dpi: int = 144,
        timeout: int = 180,
        per_sheet: bool = True,
    ) -> dict[str, list[Path]] | list[Path]:
        """Render the workbook to PNG using LibreOffice headless.

        Works on Linux, macOS and Windows. The optional renderer requires
        LibreOffice and ``pdftoppm`` from Poppler; both are found on ``PATH`` or
        in their standard install directory, since the Windows and macOS
        installers do not extend ``PATH``. Use :attr:`workbook_context` when
        only the non-rendered context is needed.

        By default each sheet is rendered whole, onto one image, and the result
        is a ``{sheet name: [png]}`` mapping — the shape :meth:`to_html` shows
        under each sheet in its Sheets tab:

            >>> result.save_html(                          # doctest: +SKIP
            ...     "report.html",
            ...     screenshots=result.save_screenshots("shots/"),
            ... )

        ``per_sheet=False`` returns the flat ``list[Path]`` of print pages
        instead, laid out by the workbook's own page setup; the report then
        shows them in a separate tab, since no page can be tied to a sheet.
        A mapping is also downgraded to that flat list when the renderer does
        not produce exactly one page per sheet, rather than filing images under
        sheets they may not belong to.
        """
        from linexcel.insights import render_workbook_screenshots

        return render_workbook_screenshots(
            self._source_bytes(),
            self._source_filename,
            output_dir,
            dpi=dpi,
            timeout=timeout,
            per_sheet=per_sheet,
        )

    # -- AI documentation (optional) --------------------------------------
    def document(
        self,
        node_ids: list[str] | None = None,
        *,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
        provider: ProviderLike | None = None,
        language: str = "en",
        max_workers: int = 4,
        max_tokens: int | None = None,
        token_budget: int | None = None,
    ) -> dict[str, str]:
        """Document nodes via AI from the deterministic lineage.

        Without ``node_ids``, documents all calculation nodes
        (cells, groups, VBA).

        Provider resolution (first match wins, no implicit default and no
        preferred vendor):

        1. ``provider`` — custom LLMProvider instance or callable
        2. ``base_url`` + ``model`` (or ``LINEXCEL_AI_BASE_URL`` +
           ``LINEXCEL_AI_MODEL``) — any OpenAI-compatible endpoint, local or
           hosted

        ``language`` selects the system prompt; see :data:`linexcel.i18n.LANGUAGES`.
        ``max_workers`` caps the number of concurrent requests. Nodes that fail
        are skipped with a :class:`UserWarning`; the cards that succeeded are
        still returned.

        ``max_tokens`` caps output per node (approximate; provider-dependent).

        ``token_budget`` caps what the whole documentation run may cost, input
        and output tokens together. It is counted against :attr:`token_usage`,
        which spans the result's lifetime, so one ceiling covers this call and
        every earlier one — the figure to set is the total you are willing to
        pay for this workbook, not a per-node allowance. Nodes still queued when
        the budget is reached are left undocumented with a :class:`UserWarning`
        rather than silently billed; requests already in flight are allowed to
        finish, so treat the ceiling as approximate.

            >>> docs = result.document(base_url=..., token_budget=200_000)
            ... # doctest: +SKIP

        Tokens consumed are added to :attr:`token_usage`.
        """
        from linexcel.aidoc import document_nodes

        if node_ids is None:
            node_ids = [
                n["id"] for n in self.nodes if n.get("kind") in ("cell", "group", "vba")
            ]
        return document_nodes(
            self.graph,
            node_ids,
            model=model,
            api_key=api_key,
            base_url=base_url,
            provider=provider,
            language=language,
            max_workers=max_workers,
            usage=self.token_usage,
            max_tokens=max_tokens,
            token_budget=token_budget,
        )

    def document_workbook(
        self,
        *,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
        provider: ProviderLike | None = None,
        language: str = "en",
        max_tokens: int | None = None,
        token_budget: int | None = None,
        include_context: bool = True,
    ) -> str:
        """Document the workbook structure and calculation flow via AI.

        The response is grounded in workbook-level deterministic lineage data.
        Pass it to :meth:`to_html` or :meth:`save_html` as ``workbook_doc`` to
        display it in the viewer's separate overview tab.

        ``include_context`` adds :attr:`workbook_context` to the dossier: the
        sheet previews, cell comments, merged ranges, frozen panes and hidden
        columns — the same cues the sheet screenshots show. Without it the model
        sees how the workbook computes but not what it looks like, and describes
        a graph rather than a file. Set it to ``False`` to keep cell contents
        local; the lineage (formulas and their values) is sent either way.

        Provider resolution is the same as :meth:`document`, and tokens
        consumed are added to :attr:`token_usage`. ``max_tokens`` caps the
        output length (approximate; provider-dependent), while ``token_budget``
        caps cumulative spend on this result and raises :class:`AiDocError` if
        earlier calls already exhausted it.
        """
        from linexcel.aidoc import document_workbook

        # Context needs the workbook bytes; a result rebuilt from a graph alone
        # has none, and a structural overview is better than an exception.
        context = (
            self.workbook_context
            if include_context and self._source_data is not None
            else None
        )
        return document_workbook(
            self.graph,
            model=model,
            api_key=api_key,
            base_url=base_url,
            provider=provider,
            language=language,
            usage=self.token_usage,
            max_tokens=max_tokens,
            token_budget=token_budget,
            context=context,
        )

    def describe_screenshots(
        self,
        screenshots: Screenshots,
        *,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
        provider: ProviderLike | None = None,
        language: str = "en",
        max_tokens: int | None = None,
        token_budget: int | None = None,
    ) -> dict[str, str]:
        """Describe each sheet screenshot with a multimodal model.

        Everything else linexcel documents is grounded in the graph. A
        screenshot is the exception: colour conventions — blue inputs against
        black formulas — conditional formatting, charts and the shape of a
        layout are invisible to a text dossier however complete it is, and a
        model looking at the picture is the only way to put them in words.

        Takes what :meth:`save_screenshots` returns and gives back
        ``{sheet name: markdown}``, ready for :meth:`save_html` as
        ``screenshot_docs``. A flat list of print pages is described page by
        page, keyed by file name, since no page belongs to a single sheet.

        ``model=`` names the model that looks at the images, which is where a
        vision model is named when it differs from the writing one; provider
        resolution is otherwise :meth:`document`'s. A text-only endpoint
        raises :class:`~linexcel.aidoc.AiDocError` rather than dropping the
        picture from the request. Tokens are added to :attr:`token_usage`.
        """
        from linexcel.aidoc import describe_images

        if isinstance(screenshots, Mapping):
            by_sheet = cast("Mapping[str, Sequence[str | Path]]", screenshots)
            paired = ((name, _first_shot(shots)) for name, shots in by_sheet.items())
            images = {name: shot for name, shot in paired if shot is not None}
        else:
            images = {Path(shot).stem: shot for shot in screenshots}
        return describe_images(
            images,
            model=model,
            api_key=api_key,
            base_url=base_url,
            provider=provider,
            language=language,
            usage=self.token_usage,
            max_tokens=max_tokens,
            token_budget=token_budget,
        )

    # -- visualization -----------------------------------------------------
    def to_html(
        self,
        *,
        title: str | None = None,
        full_document: bool = True,
        docs: dict[str, str] | None = None,
        workbook_doc: str | None = None,
        screenshots: Screenshots | None = None,
        screenshot_docs: dict[str, str] | None = None,
        language: str = "en",
    ) -> str:
        """Standalone HTML document (Cytoscape) — openable in a browser.

        If ``docs`` is provided (from :meth:`document`), AI documentation
        for each node is embedded in the detail panel. If ``workbook_doc`` is
        provided (from :meth:`document_workbook`), it is shown in a separate
        overview tab. If ``screenshots`` is provided (paths or base64), they
        are displayed in a preview tab, each under the description
        ``screenshot_docs`` carries for it (from :meth:`describe_screenshots`).
        """
        graph = self.graph
        meta = dict(graph.get("meta", {}))
        # Result built without the source bytes: drop the sheet-preview tab
        # rather than failing the whole export. Tested on the attribute rather
        # than by catching RuntimeError, which would also swallow a genuine
        # extraction failure (WorkbookRenderError subclasses it).
        meta["workbookContext"] = (
            self.workbook_context if self._source_data is not None else None
        )

        if workbook_doc:
            meta["workbookDoc"] = workbook_doc

        if screenshots:
            import base64

            def _embed(s: str | Path) -> str:
                p = Path(s) if isinstance(s, (str, Path)) else None
                if p and p.exists() and p.suffix.lower() == ".png":
                    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
                    return f"data:image/png;base64,{b64}"
                return str(s)

            if isinstance(screenshots, Mapping):
                # cast: isinstance() alone leaves the Sequence arm of the union
                # in play, which erases the value type.
                by_sheet = cast("Mapping[str, Sequence[str | Path]]", screenshots)
                meta["screenshots"] = {
                    name: [_embed(s) for s in s_list]
                    for name, s_list in by_sheet.items()
                }
            else:
                meta["screenshots"] = [_embed(s) for s in screenshots]

        if screenshot_docs:
            meta["screenshotDocs"] = dict(screenshot_docs)

        graph = {
            **graph,
            "meta": meta,
            "nodes": [
                {**n, "doc": docs.get(n["id"], "") if docs else ""}
                for n in graph["nodes"]
            ],
        }
        return render_html(
            graph,
            title=title or self._title(),
            full_document=full_document,
            language=language,
        )

    def save_html(
        self,
        path: str | Path,
        *,
        title: str | None = None,
        docs: dict[str, str] | None = None,
        workbook_doc: str | None = None,
        screenshots: Screenshots | None = None,
        screenshot_docs: dict[str, str] | None = None,
        language: str = "en",
    ) -> Path:
        path = path if isinstance(path, Path) else Path(path)
        path.write_text(
            self.to_html(
                title=title,
                docs=docs,
                workbook_doc=workbook_doc,
                screenshots=screenshots,
                screenshot_docs=screenshot_docs,
                language=language,
            ),
            encoding="utf-8",
        )
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
