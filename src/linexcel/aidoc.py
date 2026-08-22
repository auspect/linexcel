"""AI-generated documentation for Excel calculations.

Vendor-neutral by construction: no provider is named in the code and none is
chosen for you. There are exactly two ways in:

- ``base_url=`` — any OpenAI-compatible endpoint (a local Ollama, vLLM or
  LM Studio runtime; a hosted gateway such as OpenRouter; OpenAI itself;
  anything else that speaks ``/chat/completions``)
- ``provider=`` — your own callable or :class:`LLMProvider` object, for an API
  that speaks something else entirely

The model doesn't guess: each node is presented with its deterministic dossier
from the graph (exact formula, step-by-step evaluation, precedents and their
values, dependents, stretched group extent, VBA links). The system prompt
enforces citing only these facts, making the documentation "provable": every
claim traces back to a formula or a workbook value.
"""

from __future__ import annotations

import base64
import json
import os
import re
import warnings
from collections.abc import Callable, Mapping
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol, runtime_checkable

from linexcel.i18n import LANGUAGES as _LANGUAGES

# No DEFAULT_MODEL: naming one would make a vendor's model the implicit choice,
# and a name that suits a hosted API is wrong for a local runtime. The model is
# always supplied by the caller, via model= or an environment variable.
MAX_DOSSIER_CHARS = 6_000
# Raised from 12k when the workbook dossier gained the presentation context
# (sheet previews and comments); _fit_workbook_dossier() sheds that part first
# if a workbook still exceeds it, so a graph-only dossier is unaffected.
MAX_WORKBOOK_DOSSIER_CHARS = 16_000
MAX_DOSSIER_COMMENT_CHARS = 300
#: Ceiling on one screenshot handed to a vision model, before base64 (which
#: adds a third). A whole sheet rendered at 144 dpi lands well under it; a
#: file above it is refused by name rather than silently posted.
MAX_IMAGE_BYTES = 8 * 1024 * 1024
#: Suffix → media type, for the data URI a vision call carries.
IMAGE_MEDIA_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif": "image/gif",
}

_PROMPTS = Path(__file__).parent / "assets" / "prompts"


def _prompts(kind: str) -> dict[str, str]:
    """Every language's ``kind`` prompt, read from ``assets/prompts/<lang>/``.

    They are prose, one document per language, and they lived as 380 lines of
    triple-quoted strings in this module: correcting a phrasing in Japanese
    meant a diff in a 1,400-line file, and adding a language meant editing
    Python and cutting a release. As files they are what they are — text a
    translator can open — and a missing one is a missing language, which the
    key-parity test catches.
    """
    found = {
        path.parent.name: path.read_text(encoding="utf-8").strip()
        for path in sorted(_PROMPTS.glob(f"*/{kind}.md"))
    }
    if not found:
        # A build that dropped the prompt files would otherwise fail much
        # later, as a KeyError on a language code, in the middle of a request.
        raise RuntimeError(
            f"No {kind} prompts under {_PROMPTS}. The package is installed "
            f"without its assets; reinstall linexcel."
        )
    return found


_SYSTEM = _prompts("node")

_WORKBOOK_SYSTEM = _prompts("workbook")

# The one prompt whose evidence is not the dossier. A screenshot carries what
# no extraction reaches — colour conventions, conditional formatting, charts,
# the shape of a layout — so the rule here is the mirror of the others: say
# what is visible, and nothing that would have to be computed to be known.
_VISION_SYSTEM = _prompts("vision")


class AiDocError(RuntimeError):
    pass


# ──────────────────────────────────────────────
# Token accounting
# ──────────────────────────────────────────────

# Runs of CJK ideographs, kana and hangul, which tokenizers split at roughly one
# token per character while ``\w+`` would swallow a whole sentence as one word.
_CJK_RE = re.compile(r"[぀-ヿ㐀-䶿一-鿿豈-﫿가-힯]")
_WORD_RE = re.compile(r"\w+")


def estimate_tokens(text: str) -> int:
    """Approximate the token count of ``text``.

    Only a fallback: :class:`TokenUsage` prefers the counts the provider
    reports. Latin script is counted as words × 4/3 (the usual 1 token ≈ 0.75
    words ratio); CJK characters are counted individually, because a Japanese
    or Chinese sentence carries no spaces and would otherwise register as a
    single word.

    >>> estimate_tokens("the quick brown fox jumps")
    6
    >>> estimate_tokens("")
    0
    """
    if not text:
        return 0
    cjk = len(_CJK_RE.findall(text))
    latin_words = len(_WORD_RE.findall(_CJK_RE.sub(" ", text)))
    return cjk + latin_words * 4 // 3


@dataclass
class TokenUsage:
    """Tokens consumed by one or more documentation requests.

    ``estimated`` is ``True`` as soon as any request in the tally had to be
    approximated by :func:`estimate_tokens` instead of being reported by the
    provider — treat such a total as an order of magnitude, not a bill.
    """

    input_tokens: int = 0
    output_tokens: int = 0
    requests: int = 0
    estimated: bool = False
    model: str = ""
    provider: str = ""

    @property
    def total(self) -> int:
        return self.input_tokens + self.output_tokens

    def add(self, other: TokenUsage) -> None:
        """Accumulate ``other`` in place, keeping the model/provider labels."""
        self.input_tokens += other.input_tokens
        self.output_tokens += other.output_tokens
        self.requests += other.requests
        self.estimated = self.estimated or other.estimated
        self.model = self.model or other.model
        self.provider = self.provider or other.provider

    def __str__(self) -> str:
        about = "~" if self.estimated else ""
        where = f" [{self.provider}/{self.model}]" if self.provider else ""
        return (
            f"{about}{self.total:,} tokens "
            f"({about}{self.input_tokens:,} in + {about}{self.output_tokens:,} out) "
            f"over {self.requests} request(s){where}"
        )


def _check_budget(token_budget: int | None, usage: TokenUsage | None) -> None:
    """Validate ``token_budget`` and report what has already been spent.

    A budget is a ceiling on *cumulative* spend, so it is compared against the
    accumulator the caller passes: documenting a workbook in several calls that
    share one :class:`TokenUsage` shares one ceiling.
    """
    if token_budget is None:
        return
    if token_budget <= 0:
        raise ValueError(f"token_budget must be > 0, got {token_budget}")
    spent = usage.total if usage is not None else 0
    if spent >= token_budget:
        raise AiDocError(
            f"Token budget of {token_budget:,} tokens is already spent "
            f"({spent:,} used by earlier calls on this result); no request was "
            "sent. Raise token_budget to continue."
        )


# ──────────────────────────────────────────────
# Provider abstraction
# ──────────────────────────────────────────────


@runtime_checkable
class LLMProvider(Protocol):
    """Minimal protocol: system + user prompt → text response."""

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str: ...


@runtime_checkable
class UsageReportingProvider(Protocol):
    """A provider that also reports what the call consumed.

    Optional: the built-in OpenAI-compatible client implements it so that token
    counts come from the API rather than from an approximation. Custom
    providers only need :class:`LLMProvider`.
    """

    def generate_with_usage(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]: ...


@runtime_checkable
class VisionProvider(Protocol):
    """A provider that can be handed an image alongside the prompts.

    Optional, and separate from :class:`LLMProvider` on purpose: most models
    served behind an OpenAI-compatible endpoint are text-only, and a caller
    asking for screenshot descriptions should be told so rather than have an
    image quietly dropped from the request.
    """

    def generate_with_image(
        self,
        system_prompt: str,
        user_prompt: str,
        image: bytes,
        *,
        media_type: str = "image/png",
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]: ...


def _generate(
    llm: LLMProvider,
    system_prompt: str,
    user_prompt: str,
    max_tokens: int | None = None,
) -> tuple[str, TokenUsage]:
    """Call ``llm``, returning its text and what the call cost.

    Providers reporting real usage are preferred; anything else is estimated.
    """
    if isinstance(llm, UsageReportingProvider):
        return llm.generate_with_usage(
            system_prompt, user_prompt, temperature=0.2, max_tokens=max_tokens
        )
    text = llm.generate(
        system_prompt, user_prompt, temperature=0.2, max_tokens=max_tokens
    )
    return text, TokenUsage(
        input_tokens=estimate_tokens(system_prompt) + estimate_tokens(user_prompt),
        output_tokens=estimate_tokens(text or ""),
        requests=1,
        estimated=True,
    )


#: Anything accepted as ``provider=``: an :class:`LLMProvider` (any object with
#: a ``generate`` method) or a plain callable with the same signature.
ProviderLike = LLMProvider | Callable[..., str]


class _CallableProvider:
    """Adapt a plain callable to the :class:`LLMProvider` protocol."""

    def __init__(self, fn: Callable[..., str]):
        self._fn = fn

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str:
        return self._fn(system_prompt, user_prompt, temperature=temperature)


def _as_provider(provider: ProviderLike) -> LLMProvider:
    """Normalize a user-supplied provider into an :class:`LLMProvider`."""
    if isinstance(provider, LLMProvider):
        return provider
    if callable(provider):
        return _CallableProvider(provider)
    raise AiDocError(
        "provider must expose a generate(system_prompt, user_prompt, *, "
        f"temperature) method or be callable, got {type(provider).__name__}"
    )


def _resolve_provider(
    *,
    provider: ProviderLike | None = None,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
) -> LLMProvider:
    """Resolve which provider to use.

    Providers are chosen explicitly; no vendor is preferred and none is
    implicit:

    1. `provider` — custom callable or LLMProvider instance
    2. `base_url` set (param or ``LINEXCEL_AI_BASE_URL`` / ``OPENAI_BASE_URL``)
       → OpenAI-compatible client, whatever sits behind that URL

    Anything else raises :class:`AiDocError` rather than picking for you.
    """
    if provider is not None:
        return _as_provider(provider)

    base = base_url or os.getenv("LINEXCEL_AI_BASE_URL") or os.getenv("OPENAI_BASE_URL")
    if base:
        resolved_model = (
            model or os.getenv("LINEXCEL_AI_MODEL") or os.getenv("OPENAI_MODEL")
        )
        if not resolved_model:
            raise AiDocError(
                f"No model named for the endpoint at {base}: pass model= or set "
                "LINEXCEL_AI_MODEL. Endpoints do not agree on a default and "
                "linexcel does not invent one."
            )
        return _OpenAICompatProvider(
            base_url=base, api_key=api_key, model=resolved_model
        )

    raise AiDocError(
        "No AI provider selected: pass base_url= with model= for any "
        "OpenAI-compatible endpoint (a local Ollama or vLLM runtime, a hosted "
        "gateway, a vendor API), or provider= for a custom LLMProvider or "
        "callable. Equivalent env vars: LINEXCEL_AI_BASE_URL and "
        "LINEXCEL_AI_MODEL. No provider is chosen implicitly."
    )


class _OpenAICompatProvider:
    """A client for anything exposing an OpenAI-compatible chat API.

    The wire format is the only thing assumed — which vendor, gateway or local
    runtime answers at ``base_url`` is the caller's business.
    """

    def __init__(self, *, base_url: str, api_key: str | None = None, model: str):
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover
            raise AiDocError(
                "openai is not installed (pip install 'linexcel[ai]' or "
                "pip install openai). It is the client for every "
                "OpenAI-compatible endpoint, not a choice of vendor."
            ) from exc
        # Local runtimes ignore the key but the client refuses to start without
        # one, so a placeholder stands in rather than a hosted key being needed.
        self._client = OpenAI(
            api_key=(
                api_key
                or os.getenv("LINEXCEL_AI_API_KEY")
                or os.getenv("OPENAI_API_KEY")
                or "not-needed"
            ),
            base_url=base_url,
        )
        self._model = model

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> str:
        return self.generate_with_usage(
            system_prompt, user_prompt, temperature=temperature, max_tokens=max_tokens
        )[0]

    def generate_with_usage(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]:
        kwargs: dict[str, Any] = {"temperature": temperature}
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                **kwargs,
            )
            text = (response.choices[0].message.content or "").strip()
        except Exception as exc:
            raise AiDocError(f"OpenAI-compatible API call failed: {exc}") from exc
        return text, _usage_from(
            getattr(response, "usage", None),
            ("prompt_tokens", "completion_tokens"),
            system_prompt + "\n\n" + user_prompt,
            text,
            model=self._model,
            provider="openai-compatible",
        )

    def generate_with_image(
        self,
        system_prompt: str,
        user_prompt: str,
        image: bytes,
        *,
        media_type: str = "image/png",
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]:
        """The same call, with the image inlined as a data URI.

        No upload and no file id: the picture travels in the request body, so
        nothing is left behind on the endpoint's side.
        """
        encoded = base64.b64encode(image).decode("ascii")
        kwargs: dict[str, Any] = {"temperature": temperature}
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": user_prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:{media_type};base64,{encoded}"
                                },
                            },
                        ],
                    },
                ],
                **kwargs,
            )
            text = (response.choices[0].message.content or "").strip()
        except Exception as exc:
            raise AiDocError(
                f"OpenAI-compatible vision call failed: {exc}. A text-only "
                "model refuses an image; name a multimodal one."
            ) from exc
        # The fallback estimate counts the prompts only — an image is worth
        # hundreds of tokens that no character count can see — so a run whose
        # endpoint reports nothing is under-counted, and says it is estimated.
        return text, _usage_from(
            getattr(response, "usage", None),
            ("prompt_tokens", "completion_tokens"),
            system_prompt + "\n\n" + user_prompt,
            text,
            model=self._model,
            provider="openai-compatible",
        )


def _usage_from(
    reported: Any,
    fields: tuple[str, str],
    prompt: str,
    text: str,
    *,
    model: str,
    provider: str,
) -> TokenUsage:
    """Build a :class:`TokenUsage` from what the API reported, else estimate.

    Local runtimes behind an OpenAI-compatible endpoint do not always populate
    the usage block, so each field falls back independently.
    """
    prompt_field, completion_field = fields
    reported_in = getattr(reported, prompt_field, None) if reported else None
    reported_out = getattr(reported, completion_field, None) if reported else None
    estimated = reported_in is None or reported_out is None
    return TokenUsage(
        input_tokens=estimate_tokens(prompt) if reported_in is None else reported_in,
        output_tokens=estimate_tokens(text) if reported_out is None else reported_out,
        requests=1,
        estimated=estimated,
        model=model,
        provider=provider,
    )


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
        "value": step.get("value") if step.get("evaluated") else "not evaluated",
    }
    if step.get("inputs"):
        out["inputs"] = step["inputs"]
    children = [_compact_steps(c) for c in step.get("children", [])]
    if children:
        out["sub_steps"] = children
    return out


def build_workbook_dossier(
    graph: dict[str, Any], *, context: dict[str, Any] | None = None
) -> dict[str, Any]:
    """Return a compact, deterministic dossier for a whole-workbook overview.

    ``context`` is a :attr:`linexcel.LineageResult.workbook_context` mapping.
    The graph alone describes how a workbook *computes*; it says nothing about
    what a reader sees on opening it — titles sitting above a table, the labels
    in the first column, cell comments, hidden columns, frozen panes. Those cues
    are exactly what the sheet screenshots show, and merging them into each
    sheet entry lets a text-only model describe the file as it looks without any
    image ever leaving the machine.

    Both parts stay deterministic: every value is read from the workbook, so the
    "cite only the dossier" rule of the system prompt still holds.
    """
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
    if context:
        sheets = _merge_presentation(sheets, context)
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


def _merge_presentation(
    sheets: list[dict[str, Any]], context: dict[str, Any]
) -> list[dict[str, Any]]:
    """Enrich each sheet entry with what a reader sees, keyed by sheet name.

    Sheets the analysis never reached — a data-only tab holds no formula, so it
    contributes no lineage node — are appended rather than dropped: a workbook
    overview that omits its input sheets misreads the file.
    """
    by_name = {sheet["name"]: sheet for sheet in sheets if sheet.get("name")}
    merged = list(sheets)
    for ctx_sheet in context.get("sheets", []):
        name = ctx_sheet.get("name")
        if name is None:
            continue
        target = by_name.get(name)
        if target is None:
            target = {"name": name, "formula_cells": 0, "lineage_nodes": {}}
            merged.append(target)
        target.update(_presentation_of(ctx_sheet))
    return merged


def _presentation_of(ctx_sheet: dict[str, Any]) -> dict[str, Any]:
    """The non-empty presentation cues of one sheet, omitting what is absent."""
    out: dict[str, Any] = {}
    if ctx_sheet.get("visibility") and ctx_sheet["visibility"] != "visible":
        out["visibility"] = ctx_sheet["visibility"]
    for key in ("freeze_panes", "hidden_columns", "merged_ranges"):
        if ctx_sheet.get(key):
            out[key] = ctx_sheet[key]
    comments = ctx_sheet.get("comments") or []
    if comments:
        out["comments"] = [
            {
                "cell": comment.get("cell"),
                "author": comment.get("author"),
                "text": (comment.get("text") or "").strip()[:MAX_DOSSIER_COMMENT_CHARS],
            }
            for comment in comments
        ]
    preview = _compact_preview(ctx_sheet.get("preview") or [])
    if preview:
        out["preview_range"] = ctx_sheet.get("preview_range")
        out["preview"] = preview
    return out


def _fit_workbook_dossier(dossier: dict[str, Any]) -> str:
    """Serialize the dossier, shedding detail until it fits the char budget.

    Ordered cheapest-loss first: long previews shrink, then the tail of the
    pattern and VBA lists, and only then are previews and comments dropped
    outright. A workbook small enough to fit loses nothing.
    """
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    for sheet in dossier["sheets"]:
        if sheet.get("preview"):
            sheet["preview"] = sheet["preview"][:4]
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    dossier["formula_patterns"] = dossier["formula_patterns"][:5]
    dossier["vba_procedures"] = dossier["vba_procedures"][:10]
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    for sheet in dossier["sheets"]:
        sheet.pop("preview", None)
        sheet.pop("preview_range", None)
        sheet.pop("comments", None)
    return json.dumps(dossier, ensure_ascii=False, default=str)


def _compact_preview(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Drop empty rows and trailing empty cells from a sheet preview.

    A preview is a fixed rectangle read from the top-left corner, so it is
    mostly padding on a sheet whose table starts lower or further right. Sending
    that padding costs tokens and tells the model nothing.
    """
    compact = []
    for row in rows:
        values = list(row.get("values") or [])
        while values and values[-1] in (None, ""):
            values.pop()
        if values:
            compact.append({"row": row.get("row"), "values": values})
    return compact


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────


def document_workbook(
    graph: dict[str, Any],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
    token_budget: int | None = None,
    context: dict[str, Any] | None = None,
) -> str:
    """Generate a Markdown overview grounded in the workbook dossier.

    Provider resolution (first match wins; no implicit default):
    1. `provider` — custom LLMProvider instance or callable
    2. `base_url` + `model` (or `LINEXCEL_AI_BASE_URL` + `LINEXCEL_AI_MODEL`) —
       any OpenAI-compatible endpoint

    ``context`` is the workbook presentation context — the sheet previews,
    comments, merged cells, frozen panes and hidden columns a reader sees when
    opening the file. Pass it to describe the workbook as it looks, not only as
    it computes; see :func:`build_workbook_dossier`.

    If a :class:`TokenUsage` is passed as ``usage``, what the call consumed is
    accumulated into it. ``token_budget`` caps cumulative spend across that
    accumulator: an already-exhausted budget raises before anything is sent.
    """
    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    _check_budget(token_budget, usage)
    dossier = build_workbook_dossier(graph, context=context)
    blob = _fit_workbook_dossier(dossier)
    llm = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
    system = _WORKBOOK_SYSTEM[language]
    user = "Workbook dossier (deterministic, extracted from workbook):\n" + blob
    try:
        text, call_usage = _generate(llm, system, user, max_tokens=max_tokens)
    except AiDocError:
        raise
    except Exception as exc:
        raise AiDocError(f"AI documentation failed: {exc}") from exc
    if usage is not None:
        usage.add(call_usage)
    return text


def describe_images(
    images: Mapping[str, bytes | str | Path],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
    token_budget: int | None = None,
) -> dict[str, str]:
    """Describe rendered images with a multimodal model, ``{name: markdown}``.

    ``images`` maps a name — a sheet name, in practice — to a PNG, either as
    bytes or as a path to read. Each one is sent on its own, so a description
    is grounded in a single picture and nothing else.

    This is the one part of linexcel whose evidence is not the deterministic
    dossier: a screenshot shows what no extraction reaches — colour
    conventions, conditional formatting, charts, the shape of a layout — and
    the prompt confines the model to what is visible rather than letting it
    reason about the calculation.

    The provider must accept an image (:class:`VisionProvider`); a text-only
    one raises rather than having the picture dropped from the request. Model
    resolution is otherwise :func:`document_workbook`'s, so ``model=`` here is
    where a vision model is named when it differs from the writing one.

    Images are sent one at a time: they are large, and the local runtimes this
    is most used against serialize them anyway. An image that fails is skipped
    with a :class:`UserWarning`; :class:`AiDocError` is raised only when every
    one failed. ``token_budget`` is checked before each call.
    """
    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    _check_budget(token_budget, usage)
    if not images:
        return {}
    llm = _resolve_provider(
        provider=provider, model=model, api_key=api_key, base_url=base_url
    )
    if not isinstance(llm, VisionProvider):
        raise AiDocError(
            f"{type(llm).__name__} cannot be handed an image: describing "
            "screenshots needs a provider exposing generate_with_image("
            "system_prompt, user_prompt, image, *, media_type). The built-in "
            "OpenAI-compatible client does; point it at a multimodal model."
        )
    system = _VISION_SYSTEM[language]
    described: dict[str, str] = {}
    failed: list[str] = []
    for name, image in images.items():
        try:
            _check_budget(token_budget, usage)
        except AiDocError:
            warnings.warn(
                f"Token budget reached: {len(images) - len(described)} "
                "screenshot(s) left undescribed.",
                UserWarning,
                stacklevel=2,
            )
            break
        try:
            payload, media_type = _image_payload(name, image)
            text, call_usage = llm.generate_with_image(
                system,
                f"Sheet: {name}",
                payload,
                media_type=media_type,
                temperature=0.2,
                max_tokens=max_tokens,
            )
        except (AiDocError, OSError) as exc:
            failed.append(f"{name} ({exc})")
            continue
        if usage is not None:
            usage.add(call_usage)
        if text.strip():
            described[name] = text.strip()
    if failed and not described:
        raise AiDocError("No screenshot could be described: " + "; ".join(failed))
    if failed:
        warnings.warn(
            f"{len(failed)} screenshot(s) not described: " + "; ".join(failed),
            UserWarning,
            stacklevel=2,
        )
    return described


def _image_payload(name: str, image: bytes | str | Path) -> tuple[bytes, str]:
    """An image as ``(bytes, media type)``, whether given as data or a path."""
    if isinstance(image, bytes | bytearray):
        payload, media_type = bytes(image), "image/png"
    else:
        path = Path(image)
        payload = path.read_bytes()
        media_type = IMAGE_MEDIA_TYPES.get(path.suffix.lower(), "image/png")
    if len(payload) > MAX_IMAGE_BYTES:
        raise AiDocError(
            f"{name}: image is {len(payload) / 1e6:.1f} MB, over the "
            f"{MAX_IMAGE_BYTES / 1e6:.0f} MB ceiling; render it at a lower dpi."
        )
    return payload, media_type


def document_nodes(
    graph: dict[str, Any],
    node_ids: list[str],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    max_workers: int = 4,
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
    token_budget: int | None = None,
) -> dict[str, str]:
    """Document the requested nodes, returns {node_id: markdown}.

    Provider resolution is the same as :func:`document_workbook` (no implicit
    default; see :func:`_resolve_provider`).

    Nodes are documented concurrently (``max_workers`` in-flight requests;
    raise it if the provider's rate limits allow). Documenting a large
    workbook is a long, often billed operation, so a node that fails does not
    discard the ones that succeeded: the successful cards are returned and a
    :class:`UserWarning` reports how many nodes were dropped.
    :class:`AiDocError` is raised only when *every* node failed.

    If a :class:`TokenUsage` is passed as ``usage``, every successful call is
    accumulated into it — including those of a run that later fails, since
    tokens already spent are still billed.

    ``token_budget`` is a ceiling on the **total** tokens the run may spend,
    input and output together, counted against ``usage`` so several calls
    sharing one accumulator share one ceiling. It is enforced between requests,
    the only point at which a cost is known: nodes still queued when the tally
    reaches the budget are never sent, and a :class:`UserWarning` reports how
    many were left undocumented. Requests already in flight are allowed to
    finish, so the final tally can exceed the budget by up to ``max_workers``
    responses — set it as an order of magnitude, not to the token. Use
    ``max_tokens`` to bound each individual response instead.
    """
    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    if max_workers < 1:
        raise ValueError("max_workers must be >= 1")
    _check_budget(token_budget, usage)
    llm = _resolve_provider(
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
    if not dossiers:
        return docs

    def _doc_one(nid_blob: tuple[str, str]) -> tuple[str, str, TokenUsage]:
        nid, blob = nid_blob
        user = "Lineage dossier (deterministic, extracted from workbook):\n" + blob
        text, call_usage = _generate(llm, system, user, max_tokens=max_tokens)
        return nid, text or "(AI returned empty response)", call_usage

    # The tally drives the budget, so it must exist even when the caller wants
    # no accumulator of their own; when they do pass one, it *is* the tally.
    tally = usage if usage is not None else TokenUsage()
    failures: list[tuple[str, Exception]] = []
    queue = iter(dossiers)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures: dict[Future[tuple[str, str, TokenUsage]], str] = {}

        def _submit_next() -> bool:
            """Queue one more node unless the budget is spent or none is left."""
            if token_budget is not None and tally.total >= token_budget:
                return False
            item = next(queue, None)
            if item is None:
                return False
            futures[pool.submit(_doc_one, item)] = item[0]
            return True

        # Nodes are submitted a few at a time rather than all at once: a budget
        # can only stop work that has not been handed to the pool yet.
        for _ in range(max_workers):
            if not _submit_next():
                break
        # Accumulating here rather than inside the workers keeps `tally` free of
        # races: this loop is the single consumer of the pool's results.
        while futures:
            done, _pending = wait(list(futures), return_when=FIRST_COMPLETED)
            for fut in done:
                node_id = futures.pop(fut)
                try:
                    nid, text, call_usage = fut.result()
                except Exception as exc:
                    failures.append((node_id, exc))
                    continue
                docs[nid] = text
                tally.add(call_usage)
            for _ in range(len(done)):
                if not _submit_next():
                    break

    # Only an exhausted budget can leave the queue undrained.
    unsent = sum(1 for _ in queue) if token_budget is not None else 0
    if unsent:
        warnings.warn(
            f"Token budget of {token_budget:,} tokens reached after "
            f"{tally.total:,}: {len(docs)} of {len(dossiers)} nodes documented, "
            f"{unsent} never sent. Raise token_budget to document the rest.",
            UserWarning,
            stacklevel=2,
        )

    if failures and not docs:
        node_id, exc = failures[0]
        raise AiDocError(
            f"AI documentation failed for all {len(failures)} nodes "
            f"(first error, node {node_id}): {exc}"
        ) from exc
    if failures:
        node_id, exc = failures[0]
        warnings.warn(
            f"AI documentation failed for {len(failures)} of {len(dossiers)} "
            f"nodes; {len(docs)} cards returned. First error (node {node_id}): "
            f"{exc}",
            UserWarning,
            stacklevel=2,
        )
    return docs
