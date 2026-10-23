"""AI-generated documentation for Excel calculations.

Vendor-neutral by construction: no provider is named in the code and none is
chosen for you. There are exactly two ways in:

- ``base_url=`` — any OpenAI-compatible endpoint (a local Ollama, vLLM or
  LM Studio runtime; a hosted gateway such as OpenRouter; OpenAI itself;
  anything else that speaks ``/chat/completions``)
- ``provider=`` — your own callable or :class:`LLMProvider` object, for an API
  that speaks something else entirely

Each node is presented with deterministic evidence from the graph and source
metadata. Prompts ask the model to cite this evidence; generated claims still
require review and are not a proof of calculation correctness.
"""

from __future__ import annotations

import base64
import hashlib
import json
import math
import os
import re
import warnings
from collections.abc import Callable, Iterable, Mapping
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
from threading import local
from types import SimpleNamespace
from typing import TYPE_CHECKING, Any, Protocol, cast, runtime_checkable
from urllib.parse import quote, urlsplit

from linexcel.ai_validation import review_protocol, strip_claims, validate_documentation
from linexcel.doc_evidence import (
    formula_evidence,
    formula_text_evidence,
    relevant_names,
    three_dimensional_reference_evidence,
    value_comparison_evidence,
    vba_excerpt_evidence,
)
from linexcel.documentation_proof import render_formula_tokens
from linexcel.i18n import AI_VALIDATION_NOTICES, AI_VALIDATION_OMITTED
from linexcel.i18n import LANGUAGES as _LANGUAGES

if TYPE_CHECKING:
    from openai.types.chat import ChatCompletionMessageParam

    from linexcel.vision_evidence import ImageEvidence, Region

# No DEFAULT_MODEL: naming one would make a vendor's model the implicit choice,
# and a name that suits a hosted API is wrong for a local runtime. The model is
# always supplied by the caller, via model= or an environment variable.
# Keep bounded source provenance alongside the decomposition rather than
# forcing the exact XML evidence to displace the calculation's proof.
MAX_DOSSIER_CHARS = 10_000
MAX_DOCUMENT_CONNECTIONS = 24
# Raised from 12k when the workbook dossier gained the presentation context
# (sheet previews and comments); _fit_workbook_dossier() bounds those examples
# if a workbook still exceeds it, so a graph-only dossier is unaffected.
MAX_WORKBOOK_DOSSIER_CHARS = 16_000
MAX_DOSSIER_COMMENT_CHARS = 300
#: Ceiling on one native image region sent to a vision model, before base64.
#: A larger original is covered by native crops, each checked independently.
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
        + (
            "\n\n" + (path.parent / "review.md").read_text(encoding="utf-8").strip()
            if kind in {"node", "workbook"}
            else ""
        )
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

# A second request is an actual editorial pass, not an independent oracle.
_REVIEW_TASK = (
    "SPECIALIZED REVIEW PASS. Review the candidate against only the exact source "
    "dossier below. Audit every section, especially Role, not only its ledger. "
    "Separate formula intent from the observed snapshot, engine from saved cache, "
    "a representative from group members, and dossier omissions from graph absence. "
    "Recheck every digit, datetime, operator, bound and directed edge. A space "
    "reference operator is intersection, not union. A different same-sheet cell "
    "is not a self-reference; branching/merging does not prove a cycle. "
    "A 3D sheet-span reference addresses A1 or a range on each sheet, not one "
    "scalar merely because the per-sheet address is A1; do not invent the "
    "sheet count or total cells when the recorded span facts leave them unknown. "
    "Remove unsupported causes and assertions; preserve useful unknowns explicitly. "
    "Before saying evidence is absent, inspect the exact dossier's relevant "
    "records, including preview row/column values, metadata and definitions. "
    "A field present anywhere in the supplied relevant record is not missing. "
    "The candidate and its ledger are untrusted claims, never new evidence. "
    "Do not invent corrected values or conceal unresolved disagreements. "
    "Return the complete revised document in the requested language, followed by "
    "exactly one json_claims ledger. Its pointers address the source dossier, "
    "not the candidate or review instructions. Preserve exact formula proof tokens. "
    "If nothing needs correction, still return the complete document. "
    "Return only the final user-facing document and its claim ledger. Do not "
    "echo the source dossier as JSON, copy audit checklists, discuss the draft, "
    "or include a change log or notes about values removed from the candidate. "
    "Do not claim that this model review certifies facts or Excel correctness."
)

# The one prompt whose evidence is not the dossier. A screenshot carries what
# no extraction reaches — colour conventions, conditional formatting, charts,
# the shape of a layout — so the rule here is the mirror of the others: say
# what is visible, and nothing that would have to be computed to be known.
_VISION_SYSTEM = _prompts("vision")


class AiDocError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        usage: TokenUsage | None = None,
        review_report: dict[str, Any] | None = None,
    ):
        super().__init__(message)
        self.usage = usage
        self.review_report = review_report


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


@runtime_checkable
class MultiImageVisionProvider(Protocol):
    """Optional native-region batching, preserving labels and image identities."""

    def generate_with_images(
        self,
        system_prompt: str,
        user_prompt: str,
        images: list[tuple[str, bytes, str]],
        *,
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
                "No model named for the endpoint: pass model= or set "
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


class OpenAICompatProvider:
    """A client for anything exposing an OpenAI-compatible chat API.

    The wire format is the only thing assumed — which vendor, gateway or local
    runtime answers at ``base_url`` is the caller's business.

    ``base_url`` may contain ``{model}`` in its hostname or path. The model
    is also sent in the JSON body. Pass this object as ``provider=`` to node,
    workbook or image documentation. ``client_kwargs`` configures the SDK;
    ``http_client_kwargs`` configures its default HTTPX transport (HTTPX2 in
    newer SDKs); ``request_kwargs`` configures chat completions, including
    ``extra_body``, ``extra_headers`` and ``extra_query``.

    ``streaming=True`` explicitly assembles content deltas and terminal usage.
    A stream must finish with ``stop``; interrupted streams are never accepted
    as documentation. SDK retries do not restart an interrupted stream.

    Use a context manager to close connections created here. Injected SDK
    clients and HTTP clients remain caller-owned and are never closed here.
    """

    def __init__(
        self,
        *,
        model: str,
        base_url: str | None = None,
        api_key: str | None = None,
        proxy: str | None = None,
        http_client: Any = None,
        http_client_kwargs: Mapping[str, Any] | None = None,
        client_kwargs: Mapping[str, Any] | None = None,
        request_kwargs: Mapping[str, Any] | None = None,
        streaming: bool = False,
        client: Any = None,
    ):
        if not model:
            raise ValueError("model must be named")
        self._model = model
        self._streaming = streaming
        self._vision_state = local()
        self._request_kwargs = dict(request_kwargs or {})
        if {"model", "messages", "max_tokens"} & self._request_kwargs.keys():
            raise ValueError(
                "request_kwargs cannot override model, messages or max_tokens"
            )
        extra_body = self._request_kwargs.get("extra_body") or {}
        if {"model", "messages", "max_tokens"} & extra_body.keys():
            raise ValueError("extra_body cannot override model, messages or max_tokens")
        if "stream" in self._request_kwargs or "stream" in extra_body:
            raise ValueError(
                "Use the explicit streaming argument, not request_kwargs stream"
            )
        for options in (self._request_kwargs, extra_body):
            if "n" in options and (type(options["n"]) is not int or options["n"] != 1):
                raise ValueError("Documentation requires a single choice (n=1)")
        self._owns_client = client is None and http_client is None
        if client is not None:
            if any(
                (
                    base_url,
                    api_key,
                    proxy,
                    http_client is not None,
                    http_client_kwargs,
                    client_kwargs,
                )
            ):
                raise ValueError(
                    "client cannot be combined with SDK or HTTP configuration"
                )
            self._client = client
            return
        base = (
            base_url
            or os.getenv("LINEXCEL_AI_BASE_URL")
            or os.getenv("OPENAI_BASE_URL")
        )
        if not base:
            raise AiDocError("No AI provider selected: pass base_url= or client=")
        base = _model_base_url(base, model)
        options = dict(client_kwargs or {})
        if {"api_key", "base_url", "http_client"} & options.keys():
            raise ValueError("Use explicit api_key, base_url and http_client arguments")
        http_options = dict(http_client_kwargs or {})
        if http_client is not None and (proxy is not None or http_options):
            raise ValueError(
                "http_client cannot be combined with proxy or http_client_kwargs"
            )
        if proxy is not None:
            if "proxy" in http_options:
                raise ValueError("Specify proxy only once")
            http_options["proxy"] = proxy
        try:
            from openai import DefaultHttpxClient, OpenAI
        except ImportError as exc:  # pragma: no cover
            raise AiDocError(
                "openai is not installed (pip install 'linexcel[ai]' or "
                "pip install openai). It is the client for every "
                "OpenAI-compatible endpoint, not a choice of vendor."
            ) from exc
        # Local runtimes ignore the key but the client refuses to start without
        # one, so a placeholder stands in rather than a hosted key being needed.
        if "timeout" not in options:
            timeout = float(os.getenv("LINEXCEL_AI_TIMEOUT_SECONDS", "300"))
            if not math.isfinite(timeout) or timeout <= 0:
                raise ValueError(
                    "LINEXCEL_AI_TIMEOUT_SECONDS must be finite and positive"
                )
            options["timeout"] = timeout
        options.setdefault("max_retries", 0)
        created_http = None
        if http_options:
            created_http = DefaultHttpxClient(**http_options)
            http_client = created_http
        try:
            self._client = OpenAI(
                api_key=(
                    api_key
                    or os.getenv("LINEXCEL_AI_API_KEY")
                    or os.getenv("OPENAI_API_KEY")
                    or "not-needed"
                ),
                base_url=base,
                http_client=http_client,
                **options,
            )
        except Exception:
            if created_http is not None:
                created_http.close()
            raise

    def close(self) -> None:
        """Close owned connections after all in-flight calls have finished."""
        if self._owns_client:
            self._client.close()

    def __enter__(self) -> OpenAICompatProvider:
        return self

    def __exit__(self, *args: Any) -> None:
        self.close()

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
        kwargs: dict[str, Any] = {"temperature": temperature, **self._request_options()}
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        return self._complete(
            [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            system_prompt + "\n\n" + user_prompt,
            kwargs,
        )

    def _complete(
        self,
        messages: list[dict[str, Any]],
        prompt: str,
        options: dict[str, Any],
        *,
        vision: bool = False,
    ) -> tuple[str, TokenUsage]:
        streaming = getattr(self, "_streaming", False)
        if streaming:
            options["stream"] = True
            # Extra-body SDK fields can override top-level options; consolidate
            # both sources before forcing terminal usage for the stream.
            extra_body = dict(options.get("extra_body") or {})
            stream_options = {
                **(options.get("stream_options") or {}),
                **(extra_body.pop("stream_options", None) or {}),
                "include_usage": True,
            }
            options["stream_options"] = stream_options
            if "extra_body" in options:
                options["extra_body"] = extra_body
        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=cast("Iterable[ChatCompletionMessageParam]", messages),
                **options,
            )
            if streaming:
                return self._consume_stream(response, prompt)
            text = (response.choices[0].message.content or "").strip()
        except AiDocError:
            raise
        except Exception as exc:
            kind = "vision" if vision else "API"
            hint = (
                " Check gateway settings and the model's image support."
                if vision
                else ""
            )
            raise AiDocError(
                f"OpenAI-compatible {kind} call failed ({_api_error_name(exc)}).{hint}"
            ) from exc
        return self._checked_response(response, prompt, text)

    def _consume_stream(self, stream: Any, prompt: str) -> tuple[str, TokenUsage]:
        parts: list[str] = []
        reported = None
        finish_reason = None
        resolved_model = None
        response_id = None
        failure = None
        try:
            for chunk in stream:
                chunk_model, chunk_id = (
                    getattr(chunk, "model", None),
                    getattr(chunk, "id", None),
                )
                if chunk_model:
                    if resolved_model is not None and chunk_model != resolved_model:
                        raise AiDocError("AI stream changed its resolved model")
                    resolved_model = chunk_model
                if chunk_id:
                    if response_id is not None and chunk_id != response_id:
                        raise AiDocError("AI stream changed its response id")
                    response_id = chunk_id
                if getattr(chunk, "usage", None) is not None:
                    reported = chunk.usage
                choices = getattr(chunk, "choices", None) or []
                if len(choices) > 1:
                    raise AiDocError(
                        "AI stream contains multiple choices; expected index 0"
                    )
                for choice in choices:
                    if getattr(choice, "index", None) != 0:
                        raise AiDocError(
                            "AI stream contains an unexpected choice index; expected 0"
                        )
                    if finish_reason is not None:
                        raise AiDocError(
                            "AI stream returned choices after a terminal finish reason"
                        )
                    content = getattr(choice.delta, "content", None)
                    if content is not None:
                        if not isinstance(content, str):
                            raise AiDocError(
                                "AI stream returned a non-text content delta"
                            )
                        parts.append(content)
                    # reasoning_content and reasoning_details are not prose.
                    finish_reason = getattr(choice, "finish_reason", None)
        except AiDocError as exc:
            failure = exc
            if exc.usage is None:
                exc.usage = _usage_from(
                    reported,
                    ("prompt_tokens", "completion_tokens"),
                    prompt,
                    "".join(parts),
                    model=self._model,
                    provider="openai-compatible",
                )
            raise
        except Exception as exc:
            usage = _usage_from(
                reported,
                ("prompt_tokens", "completion_tokens"),
                prompt,
                "".join(parts),
                model=self._model,
                provider="openai-compatible",
            )
            failure = AiDocError(
                f"OpenAI-compatible stream failed ({_api_error_name(exc)})",
                usage=usage,
            )
            raise failure from exc
        finally:
            try:
                stream.close()
            except Exception as exc:
                usage = _usage_from(
                    reported,
                    ("prompt_tokens", "completion_tokens"),
                    prompt,
                    "".join(parts),
                    model=self._model,
                    provider="openai-compatible",
                )
                if failure is not None:
                    failure.add_note(
                        f"Stream close also failed ({_api_error_name(exc)})"
                    )
                else:
                    raise AiDocError(
                        "OpenAI-compatible stream close failed "
                        f"({_api_error_name(exc)})",
                        usage=usage,
                    ) from exc
        # A completed receipt follows successful close, so a subclass cannot
        # record success before a close failure makes the whole call fail.
        text = "".join(parts).strip()
        assembled = SimpleNamespace(
            model=resolved_model,
            id=response_id,
            usage=reported,
            choices=[
                SimpleNamespace(
                    index=0,
                    finish_reason=finish_reason,
                    message=SimpleNamespace(content=text),
                )
            ],
        )
        return self._checked_response(assembled, prompt, text, require_stop=True)

    def _checked_response(
        self,
        response: Any,
        prompt: str,
        text: str,
        *,
        require_stop: bool = False,
    ) -> tuple[str, TokenUsage]:
        status, usage = "failed", None
        try:
            output, usage = self._checked_output(
                getattr(response, "usage", None),
                prompt,
                text,
                getattr(response.choices[0], "finish_reason", None),
                require_stop=require_stop,
            )
            status = "complete"
            return output, usage
        except AiDocError as exc:
            usage = exc.usage
            raise
        finally:
            state = getattr(self, "_vision_state", None)
            if state is not None and getattr(state, "active", False):
                state.receipt = {
                    "status": status,
                    "requested_model": self._model,
                    "resolved_model": getattr(response, "model", None),
                    "response_id": getattr(response, "id", None),
                    "finish_reason": getattr(
                        response.choices[0], "finish_reason", None
                    ),
                    "streaming": require_stop,
                    "prompt_text_sha256": hashlib.sha256(prompt.encode()).hexdigest(),
                    "content_sha256": hashlib.sha256(text.encode()).hexdigest(),
                    "content_characters": len(text),
                    "usage": vars(usage).copy() if usage is not None else None,
                }

    def _checked_output(
        self,
        reported: Any,
        prompt: str,
        text: str,
        reason: Any,
        *,
        require_stop: bool = False,
    ) -> tuple[str, TokenUsage]:
        usage = _usage_from(
            reported,
            ("prompt_tokens", "completion_tokens"),
            prompt,
            text,
            model=self._model,
            provider="openai-compatible",
        )
        if require_stop and reason is None:
            raise AiDocError("AI stream ended without finish_reason=stop", usage=usage)
        if reason not in (None, "stop"):
            raise AiDocError(
                f"AI response is incomplete (finish_reason={reason}). "
                "Increase max_tokens for a length limit, "
                "or inspect the provider response.",
                usage=usage,
            )
        if not text:
            raise AiDocError("AI returned empty response", usage=usage)
        return text, usage

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
        return self.generate_with_images(
            system_prompt,
            user_prompt,
            [("image", image, media_type)],
            temperature=temperature,
            max_tokens=max_tokens,
        )

    def generate_with_images(
        self,
        system_prompt: str,
        user_prompt: str,
        images: list[tuple[str, bytes, str]],
        *,
        temperature: float = 0.2,
        max_tokens: int | None = None,
    ) -> tuple[str, TokenUsage]:
        """Send labelled native crops; the prompt supplies their exact manifest."""
        if not images:
            raise AiDocError("A vision request requires at least one image")
        content: list[dict[str, Any]] = [{"type": "text", "text": user_prompt}]
        for label, image, media_type in images:
            # Preserve the historic one-image request shape for direct callers.
            if len(images) > 1 or label != "image":
                content.append({"type": "text", "text": f"Region: {label}"})
            encoded = base64.b64encode(image).decode("ascii")
            content.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{media_type};base64,{encoded}",
                        "detail": "high",
                    },
                }
            )
        kwargs: dict[str, Any] = {"temperature": temperature, **self._request_options()}
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        # The fallback estimate counts the prompts only — an image is worth
        # hundreds of tokens that no character count can see — so a run whose
        # endpoint reports nothing is under-counted, and says it is estimated.
        if not hasattr(self, "_vision_state"):
            self._vision_state = local()
        self._vision_state.active = True
        self._vision_state.receipt = None
        try:
            return self._complete(
                [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": content},
                ],
                system_prompt + "\n\n" + user_prompt,
                kwargs,
                vision=True,
            )
        finally:
            self._vision_state.active = False

    @property
    def vision_receipt(self) -> dict[str, Any] | None:
        """Last image completion identity in this thread, excluding credentials."""
        return getattr(getattr(self, "_vision_state", None), "receipt", None)

    def _request_options(self) -> dict[str, Any]:
        # Older callers and response-contract tests construct the private class
        # without __init__; retain that compatibility.
        return getattr(self, "_request_kwargs", {})


def _model_base_url(base_url: str, model: str) -> str:
    """Expand the literal model placeholder without changing URL structure."""
    parts = urlsplit(base_url)
    if "{model}" in parts.netloc:
        if not re.fullmatch(r"[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?", model):
            raise ValueError(
                "A model used in a hostname must contain only DNS characters"
            )
    expanded = base_url.replace("{model}", quote(model, safe="."))
    parsed = urlsplit(expanded)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise ValueError("base_url must be an absolute HTTP(S) URL")
    if parsed.username is not None or parsed.password is not None:
        raise ValueError("Pass credentials through api_key or headers, not base_url")
    if parsed.query or parsed.fragment or "{" in expanded or "}" in expanded:
        raise ValueError(
            "base_url accepts only {model}; use default_query for query parameters"
        )
    return expanded


# Preserve the existing internal resolution and monkeypatching hook.
_OpenAICompatProvider = OpenAICompatProvider


def _api_error_name(exc: Exception) -> str:
    """Report the HTTP status without copying provider payloads or URLs."""
    status = getattr(exc, "status_code", None)
    suffix = f", HTTP {status}" if isinstance(status, int) else ""
    return type(exc).__name__ + suffix


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
    return _build_dossier(_index_dossiers(graph), node_id)


def _index_dossiers(graph: dict[str, Any]) -> tuple[dict, dict, dict, dict]:
    """Index adjacency once for a batch, instead of rescanning every edge."""
    nodes = {n["id"]: n for n in graph["nodes"]}
    incoming: dict[str, list[dict]] = {}
    outgoing: dict[str, list[dict]] = {}
    for edge in graph["edges"]:
        incoming.setdefault(edge["target"], []).append(edge)
        outgoing.setdefault(edge["source"], []).append(edge)
    return nodes, incoming, outgoing, graph.get("meta", {})


def _build_dossier(index: tuple[dict, dict, dict, dict], node_id: str) -> dict | None:
    nodes, incoming, outgoing, meta = index
    node = nodes.get(node_id)
    if node is None:
        return None
    precedents = incoming.get(node_id, [])
    dependents = outgoing.get(node_id, [])
    facts = formula_evidence(
        node.get("formula"), sheet=node.get("sheet"), address=node.get("addr")
    )
    dossier: dict[str, Any] = {
        "node_id": node_id,
        "kind": node.get("kind"),
        "sheet": node.get("sheet"),
        "address": node.get("addr"),
        "formula": node.get("formula"),
        "r1c1_form": node.get("r1c1"),
        "group_cells": node.get("count"),
        "extent": node.get("bbox"),
        "extent_is_bounding_box": node.get("kind") == "group",
        # A file cache, volatile snapshot or unknown provenance is not proof
        # of a successful recalculation. Keep its value with an explicit source.
        "computed_value": node.get("value")
        if node.get("valueSource") == "engine"
        else None,
        "recalculation_engine": meta.get("engine", "unspecified"),
        **_value_evidence(node),
        "value_samples": node.get("samples"),
        "decomposition": _compact_steps(node.get("steps")),
        "formula_facts": facts,
        "source_defined_names": _document_name_evidence(
            relevant_names(
                meta.get("definedNameEvidence", {}),
                facts.get("references", []),
                node.get("sheet"),
            )
        ),
        "direct_self_edge_observed": any(e["source"] == node_id for e in precedents),
        "precedents": [
            _neighbor(nodes.get(e["source"], {}), e) for e in precedents[:30]
        ],
        "dependents": [
            _neighbor(nodes.get(e["target"], {}), e) for e in dependents[:30]
        ],
        "neighbor_coverage": {
            "scope": "supplied_graph_edges",
            "omissions_scope": "dossier_only",
            "precedents_total": len(precedents),
            "dependents_total": len(dependents),
            "precedents_omitted": max(0, len(precedents) - 30),
            "dependents_omitted": max(0, len(dependents) - 30),
            "interpretation": (
                "Omitted neighbors are withheld from this dossier, not absent "
                "from the supplied graph. A group neighbor identifies its "
                "representative, not every member's reference."
            ),
        },
    }
    dossier["review_protocol"] = review_protocol()
    dossier.update(_formula_provenance(node))
    dossier["value_policy"] = {
        "volatile_cache_preservation": node.get("valueSource") == "volatile",
        "interpretation": (
            "volatile source preserves a saved reading by policy; "
            "it does not prove engine inability. Saved cache age and "
            "mismatch causes are unknown."
        ),
    }
    if node.get("kind") == "group":
        dossier["group_coverage"]["value_samples_in_dossier"] = len(
            node.get("samples") or []
        )
    if node.get("kind") == "vba":
        dossier["vba"] = {
            "module": node.get("module"),
            "procedure": node.get("proc"),
            "type": node.get("procKind"),
            **vba_excerpt_evidence(node),
            "lines": node.get("lines"),
            "provenance": node.get("vbaProvenance"),
            "execution": "not_executed",
        }
    if node.get("kind") == "opaque":
        source = node.get("valueSource")
        dossier["external_resolution"] = (
            "read_from_disk"
            if source == "external"
            else "embedded_file_cache"
            if source == "file"
            else "unresolved"
        )
        dossier["external_resolution_interpretation"] = (
            "read_from_disk means an external file was opened and read; "
            "embedded_file_cache means values come from an internal cache "
            "stored in this workbook; "
            "unresolved establishes no disk read or cached value. An opaque node "
            "may be an internal 3-D/structured reference; "
            "it does not prove an external workbook."
        )
    return dossier


def _formula_provenance(node: dict) -> dict:
    """Stored XML, a canonical expression and a calculation trace are distinct."""
    return {
        "source_formula": node.get("sourceFormula"),
        "canonical_formula": node.get("canonicalFormula", node.get("formula")),
        "formula_provenance": node.get("formulaProvenance", {"status": "unavailable"}),
        "source_value_provenance": node.get("sourceValueProvenance"),
        "formula_text_evidence": formula_text_evidence(node),
    }


def _document_name_evidence(evidence: dict) -> dict:
    """Keep definitions, without confusing metadata reading with evaluation."""
    return {
        **evidence,
        "definition_provenance": "source_workbook_metadata",
        "scope_counts_in_supplied_definitions": {
            scope: sum(
                item.get("scope") == scope for item in evidence.get("definitions", [])
            )
            for scope in ("workbook", "sheet", "unknown")
        },
        "definition_count_scope": (
            "all scopes; filtered/omitted definitions are not a complete inventory"
        ),
        "definitions": [
            {key: value for key, value in item.items() if key != "expression_evaluated"}
            for item in evidence.get("definitions", [])
        ],
    }


def _value_evidence(node: dict) -> dict:
    origin = node.get("valueSource", "unknown")
    if origin == "engine":
        if node.get("formula"):
            origin_kind = "formula_result"
        elif node.get("kind") == "input":
            origin_kind = "input_read"
        elif node.get("kind") == "name":
            origin_kind = "name_resolution"
        else:
            origin_kind = "engine_value_without_formula"
    elif origin == "external":
        origin_kind = "external_file_read"
    else:
        origin_kind = {
            "file": "file_cache",
            "volatile": "volatile_cache",
        }.get(origin, "unknown")
    evidence = {
        "displayed_value": node.get("value"),
        "value_source": node.get("valueSource", "unknown"),
        "value_origin_kind": origin_kind,
        "cached_value": node.get("cachedValue"),
        "cached_value_kind": node.get("cachedValueKind", "unknown"),
        "displayed_value_kind": node.get("valueKind", "unknown"),
        "cached_agreement": node.get("cachedAgreement"),
        "group_cached_agreement": node.get("groupCachedAgreement"),
        "verification": node.get("verification", "not_independently_verified"),
        "semantic_risks": node.get("semanticRisks", []),
        "value_comparison": value_comparison_evidence(node),
        "result_evidence": {
            "recorded_engine_formula_result": origin == "engine"
            and origin_kind == "formula_result"
            and node.get("value") is not None,
            "scope": (
                "this node or representative only; not independent Excel verification"
            ),
            "saved_cache_is_engine_result": False,
            "error_cause": "not_established",
        },
    }
    if node.get("kind") == "group":
        evidence["value_role"] = "representative cell, not an aggregate of the group"
        evidence["representative_cell"] = {
            "sheet": node.get("sheet"),
            "address": node.get("addr"),
        }
        evidence["group_members"] = node.get("count")
        evidence["group_coverage"] = _group_coverage(node)
    return evidence


def _range_dimensions(reference: Any) -> dict[str, int] | None:
    """A1 rectangle geometry only; no resolution of names/dynamic expressions."""
    bounds = _range_bounds(reference)
    if bounds is None:
        return None
    row1, col1, row2, col2 = bounds
    rows, columns = row2 - row1 + 1, col2 - col1 + 1
    return {"rows": rows, "columns": columns, "cells": rows * columns}


def _range_bounds(reference: Any) -> tuple[int, int, int, int] | None:
    """Only local or single-sheet A1 rectangles, never 3-D/external references."""
    if not isinstance(reference, str):
        return None
    address = reference
    if "!" in reference:
        qualified = re.fullmatch(
            r"(?:'((?:[^']|'')+)'|([\w.]+))!([A-Za-z$0-9:]+)", reference
        )
        if not qualified or any(
            char in (qualified[1] or qualified[2]) for char in ":[]"
        ):
            return None
        address = qualified[3]
    address = address.upper()
    bounds = address.split(":")
    if len(bounds) not in (1, 2):
        return None

    def endpoint(text: str) -> tuple[int | None, int | None] | None:
        match = re.fullmatch(r"(\$?[A-Z]{1,3})?(\$?[1-9][0-9]{0,6})?", text)
        if not match or not any(match.groups()):
            return None
        col = 0
        for char in (match[1] or "").lstrip("$"):
            col = col * 26 + ord(char) - ord("A") + 1
        row = int(match[2].lstrip("$")) if match[2] else None
        if col > 16384 or (row is not None and row > 1048576):
            return None
        return row, col or None

    first, last = endpoint(bounds[0]), endpoint(bounds[-1])
    if first is None or last is None:
        return None
    if len(bounds) == 1 and None in first:
        return None  # A bare column/row token can instead be a defined name.
    if (first[0] is None) != (last[0] is None) or (
        (first[1] is None) != (last[1] is None)
    ):
        return None
    rows = (last[0] or 1048576) - (first[0] or 1) + 1
    columns = (last[1] or 16384) - (first[1] or 1) + 1
    if rows <= 0 or columns <= 0:
        return None
    return first[0] or 1, first[1] or 1, last[0] or 1048576, last[1] or 16384


def _group_coverage(node: dict) -> dict:
    """Membership geometry and bounded value sampling are separate evidence."""
    bounds = _range_bounds(node.get("bbox"))
    dimensions = None
    if bounds:
        rows, columns = bounds[2] - bounds[0] + 1, bounds[3] - bounds[1] + 1
        dimensions = {"rows": rows, "columns": columns, "cells": rows * columns}
    count = node.get("count")
    valid_count = isinstance(count, int) and not isinstance(count, bool) and count > 0
    area = dimensions["cells"] if dimensions else None
    membership = "unknown"
    samples = node.get("samples")
    addresses = None
    conflicts = []

    def local_member(address: Any) -> bool:
        if not isinstance(address, str) or not re.fullmatch(
            r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*", address
        ):
            return False
        cell = _range_bounds(address)
        return bool(
            bounds
            and cell
            and bounds[0] <= cell[0] <= bounds[2]
            and bounds[1] <= cell[1] <= bounds[3]
        )

    representative = node.get("addr")
    if representative is not None and not local_member(representative):
        conflicts.append("representative_outside_or_invalid_bbox")
    if isinstance(samples, list) and all(
        isinstance(sample, dict) and local_member(sample.get("addr"))
        for sample in samples
    ):
        addresses = {sample["addr"].replace("$", "").upper() for sample in samples}
        if len(addresses) != len(samples) or (valid_count and len(addresses) > count):
            conflicts.append("inconsistent_sample_count")
    elif samples is not None:
        conflicts.append("sample_outside_or_invalid_bbox")
    if valid_count and area and count > area:
        conflicts.append("member_count_exceeds_bbox")
    if not conflicts and valid_count and area and local_member(representative):
        membership = "complete_bbox" if count == area else "partial_bbox"
    if conflicts:
        addresses = None
    representative = (
        representative.replace("$", "").upper()
        if isinstance(representative, str)
        else ""
    )
    sampled = len(addresses) if addresses is not None else None
    return {
        "basis": "graph_membership",
        "conflicts": conflicts,
        "member_cells": count,
        "bbox_dimensions": dimensions,
        "membership": membership,
        "sampled_cells_in_graph": sampled,
        "representative_in_samples": representative in addresses
        if addresses is not None
        else None,
        "other_sampled_cells_in_graph": len(addresses - {representative})
        if addresses is not None
        else None,
        "value_samples_in_dossier": 0,
    }


def _fit_node_dossier(dossier: dict) -> str:
    """Trim supporting examples before losing the calculation's proof."""

    def encoded() -> str:
        return json.dumps(dossier, ensure_ascii=False, default=str)

    if len(encoded()) <= MAX_DOSSIER_CHARS:
        return encoded()
    for limit in (10, 3, 1, 0):
        for direction in ("precedents", "dependents"):
            dossier[direction] = dossier[direction][:limit]
            coverage = dossier["neighbor_coverage"]
            coverage[f"{direction}_omitted"] = coverage[f"{direction}_total"] - len(
                dossier[direction]
            )
        if len(encoded()) <= MAX_DOSSIER_CHARS:
            return encoded()
    names = dossier.get("source_defined_names", {})
    if len(names.get("definitions", [])) > 4:
        names["omitted_for_size"] = len(names["definitions"]) - 4
        names["definitions"] = names["definitions"][:4]
    samples = dossier.get("value_samples") or []
    if samples:
        dossier["value_samples_omitted"] = len(samples)
        dossier["value_samples"] = []
        if "group_coverage" in dossier:
            dossier["group_coverage"]["value_samples_in_dossier"] = 0
    if len(encoded()) > MAX_DOSSIER_CHARS:
        dossier["decomposition"] = (
            "omitted: dossier size limit; not a calculation proof"
        )
    if len(encoded()) > MAX_DOSSIER_CHARS:
        # Never pass a sliced formula as if it were the complete calculation.
        # Oversized values/formulas carry an explicit omission instead.
        def bound(value: Any) -> Any:
            if isinstance(value, str) and len(value) > 1000:
                return {
                    "omitted": "dossier size limit",
                    "original_characters": len(value),
                }
            if isinstance(value, dict):
                return {key: bound(item) for key, item in value.items()}
            if isinstance(value, list):
                return [bound(item) for item in value]
            return value

        dossier = bound(dossier)
        vba = dossier.get("vba", {})
        if isinstance(vba.get("code"), dict) and vba["code"].get("omitted"):
            # The measured excerpt was replaced by an omission, so it is no
            # longer supplied code. Keep full/graph facts, update the sent stage.
            bounds = vba["code_bounds"]
            bounds["dossier_characters"] = 0
            bounds["omitted_graph_to_dossier"] = bounds["graph_characters"]
            bounds["omitted_full_source_to_dossier"] = (
                bounds["full_source_characters"]
                if bounds["source_and_graph_size_consistent"] is True
                else None
            )
            vba["code_omitted_characters"] = bounds["graph_characters"]
            vba["code_status"] = "omitted_by_dossier_size_limit"
            vba["code_hashes"]["dossier_sha256"] = None
    if len(encoded()) > MAX_DOSSIER_CHARS:
        raise AiDocError(
            "Node dossier exceeds size limit even after explicit omissions"
        )
    return encoded()


def _neighbor(other: dict, edge: dict) -> dict:
    out = {
        "id": other.get("id"),
        "kind": other.get("kind"),
        "label": other.get("label"),
        "edge_kind": edge.get("kind"),
        "sheet": other.get("sheet"),
        "address": other.get("addr"),
        "extent": other.get("bbox"),
        "edge_source": edge.get("source"),
        "edge_target": edge.get("target"),
        **_value_evidence(other),
        "formula": other.get("formula"),
        **_formula_provenance(other),
    }
    if other.get("kind") == "opaque":
        source = other.get("valueSource")
        out["external_resolution"] = (
            "read_from_disk"
            if source == "external"
            else "embedded_file_cache"
            if source == "file"
            else "unresolved"
        )
    return out


def _compact_steps(step: dict | None) -> dict | None:
    if step is None:
        return None
    out = {
        "expression": step.get("expr"),
        "operation": step.get("label"),
        "value": step.get("value")
        if step.get("evaluated")
        else "no evaluated value recorded",
        "evaluated": bool(step.get("evaluated")),
    }
    if not step.get("evaluated"):
        out["evaluation_attempt"] = "unknown"
    if step.get("inputs"):
        out["inputs"] = []
        for item in step["inputs"]:
            documented = dict(item)
            value = item.get("value")
            if isinstance(value, dict) and "range" in value:
                bounds = _range_bounds(value["range"])
                span = three_dimensional_reference_evidence(value["range"])
                documented["value"] = {
                    **value,
                    "n_unit": "not_determined" if span else "cells",
                    "dimensions": _range_dimensions(value["range"]),
                    **(
                        {
                            "count_scope": (
                                "reported_count_not_established_as_total_3d_cells"
                            ),
                            "three_dimensional_reference": span,
                        }
                        if span
                        else {}
                    ),
                    **(
                        {
                            "bounds": {
                                "start_row": bounds[0],
                                "start_col": bounds[1],
                                "end_row": bounds[2],
                                "end_col": bounds[3],
                            }
                        }
                        if bounds is not None
                        else {}
                    ),
                }
            out["inputs"].append(documented)
    if step.get("evaluationReason"):
        out["evaluation_reason"] = step["evaluationReason"]
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
            "formula_cells": sheet.get("formulaCells"),
            "lineage_nodes": nodes_by_sheet.get(sheet.get("name"), {}),
        }
        for sheet in sheet_stats
    ]
    if context:
        sheets = _merge_presentation(sheets, context)
    formula_patterns = sorted(
        (
            {
                "node_id": node.get("id"),
                "sheet": node.get("sheet"),
                "address": node.get("addr"),
                "formula": node.get("formula"),
                "cells": node.get("count", 1),
                "extent": node.get("bbox"),
                "formula_facts": formula_evidence(
                    node.get("formula"),
                    sheet=node.get("sheet"),
                    address=node.get("addr"),
                ),
                **_value_evidence(node),
                **_formula_provenance(node),
            }
            for node in nodes
            if node.get("kind") in {"cell", "group"}
        ),
        key=lambda item: (
            item["cells"] if isinstance(item["cells"], (int, float)) else 0
        ),
        reverse=True,
    )
    pattern_total = len(formula_patterns)
    formula_patterns = formula_patterns[:20]
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
    ext_total = stats.get("externalWorkbooks", 0)
    ext_read = stats.get("externalWorkbooksRead", 0)
    external_workbooks = {
        "workbooks_referenced": ext_total,
        "workbooks_read_from_disk": ext_read,
        "unread_workbooks": max(0, ext_total - ext_read),
        "interpretation": (
            "Workbooks not read from disk were not opened; dependent cells use "
            "embedded file caches if present, or have no value. "
            "A known cached value does not prove an external file was opened or read."
        ),
    }
    return {
        "review_protocol": review_protocol(),
        "filename": meta.get("filename"),
        # Live native handles are intentionally not returned from an isolated
        # worker. Their API availability says nothing about recalculation.
        "execution": {
            key: value
            for key, value in meta.get("execution", {"status": "unknown"}).items()
            if key in {"status", "phase", "completedPhases", "elapsedSeconds"}
        },
        "recalculation_engine": meta.get("engine", "unspecified"),
        "value_coverage": {
            "scope": meta.get("coverage", {}).get("scope", "not_provided"),
            "total_nodes": meta.get("coverage", {}).get("totalNodes"),
            "counts": {
                key: value.get("count")
                for key, value in meta.get("coverage", {}).get("categories", {}).items()
            },
            "interpretation": (
                "Counts describe graph nodes, not workbook cells. Engine formula "
                "values were recalculated in the worker; they remain independently "
                "unverified. Cache values are separate. A group represents several "
                "cells but counts once; its displayed value belongs to one member."
            ),
            "count_categories_are_value_origin_kinds": False,
        },
        "coverage": meta.get("analysisCoverage", {"status": "not_inspected"}),
        "analysis": {
            "formula_cells": stats.get("totalFormulas"),
            "count_scope": "extracted_formula_cells_not_workbook_total",
            "lineage_nodes": stats.get("totalNodes", 0),
            "lineage_edges": stats.get("totalEdges", 0),
            "grouped_patterns": stats.get("groupedPatterns", 0),
        },
        "sheets": sheets,
        "formula_patterns": formula_patterns,
        "formula_pattern_coverage": {
            "scope": "graph_formula_nodes",
            "total": pattern_total,
            "shown": len(formula_patterns),
            "omitted": pattern_total - len(formula_patterns),
        },
        "graph_connections": _document_connections(graph),
        "defined_names": defined_names,
        "source_defined_names": _document_name_evidence(
            meta.get("definedNameEvidence", {"status": "not_inspected"})
        ),
        "defined_names_in_graph_are_not_source_inventory": True,
        "vba_procedures": vba,
        "vba_source_scope": {
            "code_text": "not_supplied_in_overview",
            "execution": "not_executed",
            "interpretation": (
                "Procedure metadata is not readable full source or an executed "
                "macro trace. No omitted code can be reconstructed from names."
            ),
        },
        "external_workbooks": external_workbooks,
        "value_policy": (
            "Volatile readings marked volatile are deliberately preserved file "
            "snapshots, not engine inability; unknown values and cache mismatch "
            "causes remain unknown."
        ),
        "external_or_unresolved_references": opaque_references,
        "warnings": [
            *meta.get("warnings", []),
            *(context.get("warnings", []) if context else []),
        ],
    }


def _document_connections(graph: dict) -> dict:
    """Preserve direction and disclose the bound; do not infer a linear chain."""
    edges = graph.get("edges", [])
    connections = []
    destinations: dict[str, set[str]] = {}
    origins: dict[str, set[str]] = {}
    for edge in edges:
        source, target = edge.get("source"), edge.get("target")
        if not isinstance(source, str) or not isinstance(target, str):
            continue
        destinations.setdefault(source, set()).add(target)
        origins.setdefault(target, set()).add(source)
        if len(connections) < MAX_DOCUMENT_CONNECTIONS:
            connections.append({"source": source, "target": target})
    return {
        "scope": "supplied_graph_not_complete_workbook",
        "edges": connections,
        "total_edges": len(edges),
        "omitted_from_dossier": len(edges) - len(connections),
        "branching_observed": any(
            len(targets) > 1 for targets in destinations.values()
        ),
        "merging_observed": any(len(sources) > 1 for sources in origins.values()),
        "branching_nodes": sorted(
            source for source, targets in destinations.items() if len(targets) > 1
        )[:24],
        "merging_nodes": sorted(
            target for target, sources in origins.items() if len(sources) > 1
        )[:24],
        "degree_scope": (
            "all supplied graph edges; incoming >1 is merging, outgoing >1 is branching"
        ),
        "self_edges": [
            edge for edge in connections if edge["source"] == edge["target"]
        ],
        "self_edge_scope": (
            "only retained directed edges with identical IDs; merging/branching "
            "is not self-reference or cycle evidence"
        ),
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
            target = {
                "name": name,
                "formula_cells": None,
                "lineage_nodes": {},
                "inspection": "not_inspected",
            }
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
    outright. Retained formulas keep their complete source provenance; omitted
    examples remain counted. A workbook small enough to fit loses nothing.
    """
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    for sheet in dossier["sheets"]:
        if sheet.get("preview"):
            sheet["preview_rows_omitted_from_dossier"] = sheet.get(
                "preview_rows_omitted_from_dossier", 0
            ) + max(0, len(sheet["preview"]) - 4)
            sheet["preview"] = sheet["preview"][:4]
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    def trim_patterns(limit: int) -> None:
        dossier["formula_patterns"] = dossier["formula_patterns"][:limit]
        if "formula_pattern_coverage" in dossier:
            coverage = dossier["formula_pattern_coverage"]
            coverage["shown"] = len(dossier["formula_patterns"])
            coverage["omitted"] = coverage["total"] - coverage["shown"]

    # Formula patterns are optional examples, whereas the short VBA inventory
    # records which procedures were discovered. Reduce examples before losing
    # that inventory; every retained example still carries its source evidence.
    for limit in (5, 3, 1):
        trim_patterns(limit)
        blob = json.dumps(dossier, ensure_ascii=False, default=str)
        if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
            return blob
    dossier["vba_procedures_omitted_from_dossier"] = max(
        0, len(dossier["vba_procedures"]) - 10
    )
    dossier["vba_procedures"] = dossier["vba_procedures"][:10]
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
        return blob

    for sheet in dossier["sheets"]:
        sheet["preview_rows_omitted_from_dossier"] = sheet.get(
            "preview_rows_omitted_from_dossier", 0
        ) + len(sheet.get("preview", []))
        if sheet.get("comments"):
            sheet["comments_omitted_from_dossier"] = len(sheet["comments"])
        sheet.pop("preview", None)
        sheet.pop("preview_range", None)
        sheet.pop("comments", None)
    blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) > MAX_WORKBOOK_DOSSIER_CHARS:
        # A workbook overview can explicitly omit all optional formula examples;
        # node documentation still retains its focal source and omission rules.
        trim_patterns(0)
        blob = json.dumps(dossier, ensure_ascii=False, default=str)
    for limit in (20, 5, 0):
        if len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS:
            return blob
        for key in ("defined_names", "external_or_unresolved_references"):
            items = dossier.get(key, [])
            dossier[key + "_omitted"] = dossier.get(key + "_omitted", 0) + max(
                0, len(items) - limit
            )
            dossier[key] = items[:limit]
        names = dossier.get("source_defined_names", {})
        items = names.get("definitions", [])
        names["omitted_for_size"] = names.get("omitted_for_size", 0) + max(
            0, len(items) - limit
        )
        names["definitions"] = items[:limit]
        if "scope_counts_in_supplied_definitions" in names:
            names["scope_counts_in_supplied_definitions"] = {
                scope: sum(item.get("scope") == scope for item in names["definitions"])
                for scope in ("workbook", "sheet", "unknown")
            }
        blob = json.dumps(dossier, ensure_ascii=False, default=str)
    if len(blob) > MAX_WORKBOOK_DOSSIER_CHARS:
        raise AiDocError("Workbook dossier exceeds size limit after explicit omissions")
    return blob


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
# Structured tables
# ──────────────────────────────────────────────
#
# The model never writes Markdown pipe tables itself: a small local model
# breaks the syntax in ways a renderer cannot recover from (dropped cells,
# shifted columns, a separator row that swallows the header). Instead the
# prompts ask for a ``{{T1}}`` placeholder in the prose plus the table's
# *content* as JSON in a trailing ```json_tables block, and the functions
# below render the Markdown deterministically. Malformed JSON costs the
# tables, never the prose.

_TABLE_BLOCK_RE = re.compile(
    r"^ {0,3}```json_tables[ \t]*\r?\n(.*?)^ {0,3}```[ \t]*\r?$",
    re.DOTALL | re.IGNORECASE | re.MULTILINE,
)
_PLACEHOLDER_RE = re.compile(r"\{\{(T\d+)\}\}")
# A quantity: optional sign, digits with optional thousands spacing or US
# comma grouping, at most one decimal separator, optional %. Version strings
# ("1.2.3") and dates ("2024-01-01") deliberately fail, so they keep left
# alignment.
_NUMERIC_CELL_RE = re.compile(r"[-−+]?(\d[\d\s]*|\d{1,3}(?:,\d{3})+)([.,]\d+)?%?")


def _md_cell(value: Any) -> str:
    """One table cell, safe to drop into pipe syntax."""
    text = "" if value is None else str(value)
    return text.replace("|", "\\|").replace("\n", " ").strip()


def render_markdown_table(
    columns: list[Any],
    rows: list[list[Any]],
    *,
    caption: str | None = None,
) -> str:
    """Render a Markdown pipe table from structured data — the only place in
    linexcel where table syntax is written.

    A column whose cells all parse as numbers is right-aligned (``---:``),
    so amounts line up on their decimal separator. Rows shorter than the
    header are padded, longer ones truncated: a ragged model row can no
    longer shift every column after it.
    """
    cols = [_md_cell(c) for c in columns]
    width = len(cols)
    body = [([_md_cell(v) for v in row] + [""] * width)[:width] for row in rows]
    aligns = []
    for j in range(width):
        values = [r[j] for r in body if r[j]]
        numeric = bool(values) and all(_NUMERIC_CELL_RE.fullmatch(v) for v in values)
        aligns.append("---:" if numeric else "---")
    lines: list[str] = []
    if caption:
        lines += [f"*{caption}*", ""]
    lines.append("| " + " | ".join(cols) + " |")
    lines.append("|" + "|".join(aligns) + "|")
    lines += ["| " + " | ".join(r) + " |" for r in body]
    return "\n".join(lines)


def _unwrap_markdown(text: str) -> str:
    """Remove a model's whole-response Markdown wrapper, not real code blocks."""
    match = re.fullmatch(
        r"\s*```(?:markdown|md)[ \t]*\r?\n(.*?)\r?\n```\s*",
        text,
        re.DOTALL | re.IGNORECASE,
    )
    return match.group(1).strip() if match else text.strip()


def _insert_tables(text: str) -> str:
    """Swap ``{{Tn}}`` placeholders for code-rendered Markdown tables.

    The structured payload travels in trailing ```json_tables fenced
    blocks; the blocks themselves are removed from the returned text.
    Unparseable or incomplete payloads degrade to prose — a placeholder
    whose table is missing is dropped, never left as raw braces in the
    output, and neither is a placeholder that arrived with no block at all.
    """
    text = strip_claims(_unwrap_markdown(text))
    tables: dict[str, str] = {}
    for match in _TABLE_BLOCK_RE.finditer(text):
        try:
            payload = json.loads(match.group(1))
        except (ValueError, TypeError):
            continue
        if not isinstance(payload, list):
            continue
        for entry in payload:
            if not isinstance(entry, dict):
                continue
            tid, cols, rows = (
                entry.get("id"),
                entry.get("columns"),
                entry.get("rows"),
            )
            if (
                isinstance(tid, str)
                and isinstance(cols, list)
                and isinstance(rows, list)
                and all(isinstance(row, list) for row in rows)
            ):
                caption = entry.get("caption")
                tables[tid] = render_markdown_table(
                    cols, rows, caption=caption if isinstance(caption, str) else None
                )
    text = _TABLE_BLOCK_RE.sub("", text)
    text = _PLACEHOLDER_RE.sub(lambda m: tables.get(m.group(1), ""), text)
    # A dropped placeholder leaves a blank run behind; collapse it.
    return re.sub(r"\n{3,}", "\n\n", text).strip()


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────


def _reviewed_document(
    llm: LLMProvider,
    system: str,
    user: str,
    blob: str,
    *,
    specialized_review: bool,
    max_tokens: int | None,
    remaining_budget: int | None = None,
) -> tuple[str, TokenUsage, dict[str, Any]]:
    """Draft, then perform one bounded source-led review; preserve both phases.

    Model review remains unverified. Failed/empty review never falls back to an
    apparently reviewed draft. Calls already consumed remain in the usage tally.
    The API budget is checked between requests; a complete response already in
    flight is retained with its exact usage even if it exceeds the remainder.
    """
    tally = TokenUsage()
    audit: dict[str, Any] = {
        "enabled": specialized_review,
        "status": "pending",
        "factual_certification": False,
        "stages": [],
        "scope": "same-provider editorial review, not independent source verification",
    }
    text = ""
    for phase in ("draft", "review") if specialized_review else ("draft",):
        if phase == "review":
            draft_checks = audit["stages"][-1]["deterministic_checks"]
            user = (
                "Source dossier (only factual evidence; "
                "claim pointers address this JSON):\n"
                + blob
                + "\n\nCandidate to review (untrusted, not evidence):\n"
                + text
                + "\n\nBounded deterministic checks (not factual certification):\n"
                + json.dumps(
                    {
                        "issues": draft_checks["issues"],
                        "limitations": draft_checks["limitations"],
                        "factual_review": draft_checks["factual_review"],
                    },
                    ensure_ascii=False,
                )
            )
            system += "\n\n" + _REVIEW_TASK
        stage: dict[str, Any] = {
            "phase": phase,
            "status": "pending",
            "system_prompt": system,
            "user_prompt": user,
            "system_prompt_sha256": hashlib.sha256(system.encode()).hexdigest(),
            "user_prompt_sha256": hashlib.sha256(user.encode()).hexdigest(),
            "prompt_text_sha256": hashlib.sha256(
                (system + "\n\n" + user).encode()
            ).hexdigest(),
            "raw_response": None,
            "content_sha256": None,
            "response_id": None,
            "usage": None,
        }
        audit["stages"].append(stage)
        try:
            if remaining_budget is not None and tally.total >= remaining_budget:
                raise AiDocError("Token budget exhausted before specialized review")
            text, consumed = _generate(llm, system, user, max_tokens=max_tokens)
            tally.add(consumed)
            stage.update(
                raw_response=text,
                content_sha256=hashlib.sha256((text or "").encode()).hexdigest(),
                usage=vars(consumed).copy(),
            )
            if not text or not _insert_tables(text).strip():
                raise AiDocError("AI returned empty response")
            if len(text) > 200_000:
                raise AiDocError("AI response exceeds bounded review size")
            checks = validate_documentation(text, json.loads(blob))
            stage["deterministic_checks"] = {
                key: checks[key]
                for key in (
                    "status",
                    "issues",
                    "checks",
                    "counts",
                    "factual_review",
                    "limitations",
                )
            }
            stage["status"] = "complete"
        except Exception as exc:
            if isinstance(exc, AiDocError) and exc.usage is not None:
                tally.add(exc.usage)
                stage["usage"] = vars(exc.usage).copy()
            stage.update(status="failed", error=type(exc).__name__)
            audit["status"] = "failed"
            raise AiDocError(
                str(exc)
                if isinstance(exc, AiDocError)
                else "AI documentation request failed",
                usage=tally,
                review_report=audit,
            ) from exc
    audit["status"] = "completed" if specialized_review else "disabled"
    return text, tally, audit


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
    validation_results: dict[str, Any] | None = None,
    specialized_review: bool = True,
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
    By default, ``specialized_review`` makes two requests: draft, then review
    against the same source dossier. Both count toward usage. A budget exhausted
    after the draft prevents the review and no draft is published as reviewed.
    A final request already in flight may finish above the approximate ceiling;
    its complete response and cost are retained. ``max_tokens`` caps each
    response, not the two-request document's combined output.

    ``validation_results``, when supplied, receives the bounded quotation-check
    report under ``workbook``, including the raw response. Unsupported comparable
    quotations and processing limits add a visible qualification; ellipses and
    numerical illustrations remain explicitly unverified without alerts.
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
        text, call_usage, review = _reviewed_document(
            llm,
            system,
            user,
            blob,
            specialized_review=specialized_review,
            max_tokens=max_tokens,
            remaining_budget=None
            if token_budget is None
            else token_budget - (usage.total if usage else 0),
        )
    except AiDocError as exc:
        if usage is not None and exc.usage is not None:
            usage.add(exc.usage)
        if validation_results is not None and exc.review_report is not None:
            validation_results["workbook"] = {
                "status": "failed",
                "specialized_review": exc.review_report,
                "raw_response": None,
            }
        raise
    except Exception as exc:
        raise AiDocError(f"AI documentation failed: {exc}") from exc
    if usage is not None:
        usage.add(call_usage)
    rendered = _insert_tables(text) if text else ""
    if not rendered.strip():
        raise AiDocError("AI returned empty response")
    report = _documentation_report(text, blob)
    report["specialized_review"] = review
    rendered, insertions = render_formula_tokens(rendered, json.loads(blob))
    report["deterministic_formula_insertions"] = insertions
    if validation_results is not None:
        validation_results["workbook"] = report
    return _qualify_documentation(rendered, report, language)


def _documentation_report(text: str, blob: str) -> dict:
    """Keep exact fitted JSON independently from the parsed binding fingerprint."""
    report = validate_documentation(text, json.loads(blob))
    report["sent_dossier"] = {
        "json_text": blob,
        "json_text_sha256": hashlib.sha256(blob.encode("utf-8")).hexdigest(),
        "characters": len(blob),
        "utf8_bytes": len(blob.encode("utf-8")),
        "scope": (
            "Exact serialized JSON portion of the user prompt; excludes its "
            "prefix, the system prompt, provider request options and images."
        ),
    }
    return report


def _qualify_documentation(text: str, report: dict, language: str) -> str:
    review = report.get("specialized_review", {})
    prefix = (
        "> " + _REVIEW_NOTICES[language] + "\n\n"
        if review.get("enabled") and review.get("status") == "completed"
        else ""
    )
    if report["status"] == "qualified":
        unsupported = [
            check["quotation"]
            for check in report["checks"]
            if check["status"] == "not_source_supported"
        ]
        excerpts = []
        for number, quotation in enumerate(unsupported[:3], 1):
            # The viewer supports single-backtick spans only. Never substitute
            # characters in a formula to imitate a faithful quotation.
            if "`" in quotation:
                excerpts.append(AI_VALIDATION_OMITTED[language].format(number=number))
                continue
            excerpt = " ".join(quotation.split())
            excerpt = excerpt[:160] + "…" if len(excerpt) > 160 else excerpt
            excerpts.append(f"` {excerpt} `")
        detail = "\n>\n> " + " · ".join(excerpts) if excerpts else ""
        if len(unsupported) > 3:
            detail += f" (+{len(unsupported) - 3})"
        notice = (
            AI_VALIDATION_NOTICES[language] if unsupported else _CLAIM_NOTICES[language]
        )
        return "> " + notice + detail + "\n\n" + prefix + text
    return prefix + text


_REVIEW_NOTICES = {
    "en": (
        "AI draft and specialized review · 2 requests. Source checks are bounded; "
        "the full response is not factually certified."
    ),
    "fr": (
        "Brouillon IA puis revue spécialisée · 2 appels. "
        "Les contrôles sources sont bornés ; "
        "toute la réponse n’est pas certifiée factuellement."
    ),
    "de": (
        "KI-Entwurf und spezialisierte Prüfung · 2 Anfragen. "
        "Die Quellenprüfung ist begrenzt; die gesamte Antwort ist "
        "nicht als faktisch korrekt bestätigt."
    ),
    "es": (
        "Borrador IA y revisión especializada · 2 solicitudes. "
        "Las comprobaciones de fuentes son limitadas; "
        "la respuesta completa no está certificada como factual."
    ),
    "it": (
        "Bozza IA e revisione specialistica · 2 richieste. "
        "I controlli delle fonti sono limitati; "
        "l’intera risposta non è certificata come fattualmente corretta."
    ),
    "ja": (
        "AI草稿と専門レビュー · 2回のリクエスト。"
        "出典の確認には範囲の制限があり、"
        "回答全体の事実性は保証されません。"
    ),
    "nl": (
        "AI-concept en gespecialiseerde controle · 2 verzoeken. "
        "Broncontroles zijn begrensd; "
        "de volledige reactie is niet feitelijk gecertificeerd."
    ),
    "pt": (
        "Rascunho IA e revisão especializada · 2 pedidos. "
        "As verificações das fontes são limitadas; "
        "a resposta completa não está certificada como factual."
    ),
    "zh": (
        "AI草稿与专业复核 · 2次请求。来源检查有范围限制，并不证明整篇回答的事实准确性。"
    ),
}


_CLAIM_NOTICES = {
    "en": (
        "AI evidence review: some factual bindings are missing or disagree with "
        "the supplied evidence. Review the claims; the original response is preserved."
    ),
    "fr": (
        "Revue des preuves IA : certaines assertions factuelles manquent de preuve "
        "ou contredisent le dossier transmis. Vérifiez-les ; "
        "la réponse originale est conservée."
    ),
    "de": (
        "KI-Evidenzprüfung: Einige Faktenbindungen fehlen oder widersprechen den "
        "übermittelten Belegen. Prüfen Sie die Aussagen; "
        "die Originalantwort bleibt erhalten."
    ),
    "es": (
        "Revisión de pruebas de IA: algunas afirmaciones carecen de respaldo o "
        "contradicen el expediente enviado. Revísalas; "
        "se conserva la respuesta original."
    ),
    "it": (
        "Revisione delle prove IA: alcune affermazioni non hanno riscontri o "
        "contraddicono il dossier inviato. Verificarle; "
        "la risposta originale è conservata."
    ),
    "ja": (
        "AI証拠レビュー：一部の事実の根拠が不足しているか、送信された資料と矛盾しています。"
        "主張を確認してください。元の回答は保持されています。"
    ),
    "nl": (
        "AI-bewijscontrole: sommige feitelijke verwijzingen ontbreken of spreken "
        "het aangeleverde dossier tegen. Controleer de uitspraken; "
        "het oorspronkelijke antwoord blijft bewaard."
    ),
    "pt": (
        "Revisão de evidências da IA: algumas afirmações não têm suporte ou "
        "contradizem o dossiê enviado. Verifique-as; "
        "a resposta original é preservada."
    ),
    "zh": (
        "AI证据复核：部分事实缺少依据或与所提供的资料矛盾。"
        "请核查相关陈述；原始回答已保留。"
    ),
}


def _describe_native_batch(
    llm: VisionProvider | MultiImageVisionProvider,
    system: str,
    prompt: str,
    native: ImageEvidence,
    regions: list[Region],
    batch: dict[str, Any],
    accumulator: TokenUsage,
    *,
    token_budget: int | None,
    max_tokens: int | None,
) -> list[dict[str, Any]]:
    """One optional complete regeneration for invalid regional JSON only."""
    from linexcel.vision_evidence import VisionEvidenceError, parse_observations

    batch["attempts"] = []
    batch_usage = TokenUsage()
    for number in (1, 2):
        try:
            _check_budget(token_budget, accumulator)
        except AiDocError:
            batch["status"] = "failed"
            batch["failure_code"] = "token_budget"
            raise
        attempt_prompt = prompt
        if number == 2:
            manifest = json.loads(prompt)
            manifest["contract_retry"] = {
                "attempt": 2,
                "instruction": (
                    "Generate a complete fresh response for this same native crop. "
                    "Return only the region JSON contract, at most 3 nonempty "
                    "literal quotations of at most 80 characters. "
                    "Do not repair or shorten a previous response."
                ),
            }
            attempt_prompt = json.dumps(
                manifest, ensure_ascii=False, separators=(",", ":")
            )
        attempt: dict[str, Any] = {
            "number": number,
            "status": "pending",
            "user_prompt": attempt_prompt,
            "json_prompt_sha256": hashlib.sha256(attempt_prompt.encode()).hexdigest(),
            "prompt_text_sha256": hashlib.sha256(
                (system + "\n\n" + attempt_prompt).encode()
            ).hexdigest(),
            "system_prompt_sha256": hashlib.sha256(system.encode()).hexdigest(),
        }
        batch["attempts"].append(attempt)
        try:
            if isinstance(llm, MultiImageVisionProvider):
                text, call_usage = llm.generate_with_images(
                    system,
                    attempt_prompt,
                    [
                        (region.id, region.payload, region.media_type)
                        for region in regions
                    ],
                    temperature=0,
                    max_tokens=max_tokens,
                )
            else:
                region = regions[0]
                text, call_usage = llm.generate_with_image(
                    system,
                    attempt_prompt,
                    region.payload,
                    media_type=region.media_type,
                    temperature=0,
                    max_tokens=max_tokens,
                )
            accumulator.add(call_usage)
            batch_usage.add(call_usage)
            attempt["usage"] = vars(call_usage).copy()
            attempt["provider_receipt"] = getattr(llm, "vision_receipt", None)
            if not isinstance(text, str):
                raise AiDocError("AI returned a non-text vision response")
            attempt["raw_response"] = text
            attempt["content_sha256"] = hashlib.sha256(text.encode()).hexdigest()
            if token_budget is not None and accumulator.total > token_budget:
                attempt["failure_code"] = "token_budget"
                raise AiDocError("Vision token budget exceeded after reported usage")
            if not text.strip():
                raise AiDocError("AI returned empty response for native vision crop")
            observations = parse_observations(text, native, regions)
        except AiDocError as exc:
            if exc.usage is not None:
                accumulator.add(exc.usage)
                batch_usage.add(exc.usage)
                attempt["usage"] = vars(exc.usage).copy()
            attempt["status"] = "failed"
            attempt["provider_receipt"] = getattr(llm, "vision_receipt", None)
            raise
        except VisionEvidenceError as exc:
            attempt["status"] = "invalid_response"
            attempt["contract_error"] = str(exc)
            if number == 2:
                raise
        else:
            attempt["status"] = "complete"
            return observations
        finally:
            for key in (
                "status",
                "usage",
                "provider_receipt",
                "raw_response",
                "content_sha256",
                "contract_error",
                "failure_code",
            ):
                batch.pop(key, None)
            batch.update(
                {key: value for key, value in attempt.items() if key != "number"}
            )
            batch["final_attempt"] = number
            batch["usage_total"] = vars(batch_usage).copy()
    raise AssertionError("Native vision attempts must complete or fail")


def describe_images(
    images: Mapping[str, bytes | str | Path | Iterable[bytes | str | Path]],
    *,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    provider: ProviderLike | None = None,
    language: str = "en",
    usage: TokenUsage | None = None,
    max_tokens: int | None = None,
    token_budget: int | None = None,
    validation_results: dict[str, Any] | None = None,
    evidence_dir: str | Path | None = None,
) -> dict[str, str]:
    """Describe every native image region with bounded visual observations.

    Large screenshots are cropped without resampling into overlapping 1024px
    regions. Each request contains exactly one native crop, including for
    multi-image providers. Each response must identify the requested
    region and short contiguous quotations. Their scope is the exact
    crop computed by code; the model supplies no geometry. Rendering is
    deterministic, with no further model-generated summary. This validates
    scope, not visual truth. Exact uniform or fully transparent decoded crops
    are resolved locally: their empty text observation is not a model response
    and has no provider receipt. Nonuniform pixels always require the model.

    ``validation_results`` retains original/crop hashes, coordinates, raw
    responses and request identities, including one optional complete retry
    for invalid regional JSON. Both attempts retain their full raw response
    and consumption; no reply is truncated or repaired. ``evidence_dir`` saves
    the exact sent crops. Multiple pages belonging to one sheet are processed. A sheet
    is returned only if every page and region completes; partial evidence is
    retained, but cannot silently become a complete screenshot description.
    """
    from linexcel.vision_evidence import (
        BATCH_SIZE,
        VisionEvidenceError,
        batch_prompt,
        persist_image,
        prepare_image,
        render_observations,
        uniform_region_observations,
    )

    if language not in _LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r}. Use one of {_LANGUAGES}")
    if not images:
        return {}
    llm = None
    system = _VISION_SYSTEM[language]
    accumulator = usage if usage is not None else TokenUsage()
    described: dict[str, str] = {}
    failed: list[str] = []
    for name, image in images.items():
        report: dict[str, Any] = {
            "status": "incomplete",
            "language": language,
            "pages": [],
            "factual_certification": False,
        }
        if validation_results is not None:
            validation_results[name] = report
        try:
            sources = [image] if isinstance(image, bytes | str | Path) else list(image)
            if not sources:
                raise VisionEvidenceError("No image pages supplied")
            report["expected_pages"] = len(sources)
            for page_number, source in enumerate(sources, 1):
                payload, _ = _image_payload(name, source)
                native = prepare_image(payload)
                metadata = (
                    persist_image(native, Path(evidence_dir))
                    if evidence_dir is not None
                    else native.metadata()
                )
                page: dict[str, Any] = {
                    "evidence": metadata,
                    "batches": [],
                    "observations": [],
                    "page": page_number,
                    "source_path": (
                        str(Path(source).resolve())
                        if not isinstance(source, bytes)
                        else None
                    ),
                    "completed_regions": [],
                    "model_regions": [],
                    "deterministic_regions": [],
                }
                report["pages"].append(page)
                size = BATCH_SIZE
                for offset in range(0, len(native.regions), size):
                    regions = list(native.regions[offset : offset + size])
                    batch: dict[str, Any] = {
                        "id": f"page-{page_number}-batch-{offset // size}",
                        "region_ids": [region.id for region in regions],
                        "original_sha256": native.sha256,
                        "images": [region.metadata() for region in regions],
                        "status": "pending",
                    }
                    page["batches"].append(batch)
                    for region in regions:
                        if len(region.payload) > MAX_IMAGE_BYTES:
                            raise VisionEvidenceError(
                                f"{name}/{region.id}: crop exceeds image byte ceiling"
                            )
                    if all(r.text_visibility != "not_determined" for r in regions):
                        observations = uniform_region_observations(regions)
                        batch.update(
                            {
                                "status": "complete",
                                "resolution": "deterministic_pixels",
                                "observations": observations,
                                "attempts": [],
                                "usage_total": vars(TokenUsage()).copy(),
                                "provider_request_required": False,
                            }
                        )
                        page["observations"].extend(observations)
                        page["completed_regions"].extend(r.id for r in regions)
                        page["deterministic_regions"].extend(r.id for r in regions)
                        continue
                    _check_budget(token_budget, accumulator)
                    if llm is None:
                        llm = _resolve_provider(
                            provider=provider,
                            model=model,
                            api_key=api_key,
                            base_url=base_url,
                        )
                    if not isinstance(llm, VisionProvider | MultiImageVisionProvider):
                        raise AiDocError(
                            f"{type(llm).__name__} cannot be handed an image: "
                            "describing screenshots needs a provider exposing "
                            "generate_with_image(system_prompt, user_prompt, image, "
                            "*, media_type). The built-in "
                            "OpenAI-compatible client does; "
                            "point it at a multimodal model."
                        )
                    prompt = batch_prompt(name, page_number, native, regions)
                    batch.update(
                        {
                            "resolution": "model",
                            "provider_request_required": True,
                            "user_prompt": prompt,
                            "json_prompt_sha256": hashlib.sha256(
                                prompt.encode()
                            ).hexdigest(),
                            "prompt_text_sha256": hashlib.sha256(
                                (system + "\n\n" + prompt).encode()
                            ).hexdigest(),
                            "system_prompt_sha256": hashlib.sha256(
                                system.encode()
                            ).hexdigest(),
                        }
                    )
                    _check_budget(token_budget, accumulator)
                    observations = _describe_native_batch(
                        llm,
                        system,
                        prompt,
                        native,
                        regions,
                        batch,
                        accumulator,
                        token_budget=token_budget,
                        max_tokens=max_tokens,
                    )
                    page["observations"].extend(observations)
                    page["completed_regions"].extend(region.id for region in regions)
                    page["model_regions"].extend(region.id for region in regions)
                page["region_coverage_complete"] = set(page["completed_regions"]) == {
                    region.id for region in native.regions
                }
            report["status"] = "complete"
            described[name] = render_observations(report["pages"], language)
        except (AiDocError, OSError, VisionEvidenceError) as exc:
            report["error"] = str(exc)
            failed.append(f"{name} ({exc})")
    if failed and not described:
        raise AiDocError(
            "No screenshot could be described: " + "; ".join(failed), usage=accumulator
        )
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
    validation_results: dict[str, Any] | None = None,
    specialized_review: bool = True,
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

    ``specialized_review=True`` makes two requests per node: draft, then bounded
    editorial review by the same provider. Both requests count toward usage.
    An incomplete or unaffordable review never publishes the draft as reviewed.
    Set it to ``False`` explicitly to request only one draft per node.

    If a :class:`TokenUsage` is passed as ``usage``, consumed tokens are
    accumulated into it, including usage reported for rejected responses and
    calls in a run that later fails. Tokens already spent are still billed.

    ``token_budget`` is a ceiling on the **total** tokens the run may spend,
    input and output together, counted against ``usage`` so several calls
    sharing one accumulator share one ceiling. It is enforced between requests,
    the only point at which a cost is known: nodes still queued when the tally
    reaches the budget are never sent, and a :class:`UserWarning` reports how
    many were left undocumented. Requests already in flight are allowed to
    finish. Each active node has its own remaining-budget snapshot, so the final
    tally may include both phases of up to ``max_workers`` active nodes. Set the
    ceiling as an order of magnitude, not to the token. Use ``max_tokens`` to
    bound each individual response, not a document's combined output.

    ``validation_results``, when supplied, receives a bounded quotation-check
    report keyed by node ID, including raw responses. Formula quotations without
    source support are visibly qualified, not silently corrected. General prose
    remains unverified even when every inspected quotation has source support.
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
    index = _index_dossiers(graph)
    for nid in node_ids:
        d = _build_dossier(index, nid)
        if d is not None:
            blob = _fit_node_dossier(d)
            dossiers.append((nid, blob))
    if not dossiers:
        return docs

    def _doc_one(nid_blob: tuple[str, str]) -> tuple[str, str, TokenUsage, dict]:
        nid, blob = nid_blob
        user = "Lineage dossier (deterministic, extracted from workbook):\n" + blob
        text, call_usage, review = _reviewed_document(
            llm,
            system,
            user,
            blob,
            specialized_review=specialized_review,
            max_tokens=max_tokens,
            remaining_budget=None
            if token_budget is None
            else token_budget - tally.total,
        )
        rendered = _insert_tables(text) if text else ""
        if not rendered.strip():
            raise AiDocError("AI returned empty response", usage=call_usage)
        report = _documentation_report(text, blob)
        report["specialized_review"] = review
        rendered, insertions = render_formula_tokens(rendered, json.loads(blob))
        report["deterministic_formula_insertions"] = insertions
        return (
            nid,
            _qualify_documentation(rendered, report, language),
            call_usage,
            report,
        )

    # The tally drives the budget, so it must exist even when the caller wants
    # no accumulator of their own; when they do pass one, it *is* the tally.
    tally = usage if usage is not None else TokenUsage()
    failures: list[tuple[str, Exception]] = []
    queue = iter(dossiers)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures: dict[Future[tuple[str, str, TokenUsage, dict]], str] = {}

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
                    nid, text, call_usage, report = fut.result()
                except Exception as exc:
                    if isinstance(exc, AiDocError) and exc.usage is not None:
                        tally.add(exc.usage)
                    if (
                        validation_results is not None
                        and isinstance(exc, AiDocError)
                        and exc.review_report is not None
                    ):
                        validation_results[node_id] = {
                            "status": "failed",
                            "specialized_review": exc.review_report,
                            "raw_response": None,
                        }
                    failures.append((node_id, exc))
                    continue
                docs[nid] = text
                if validation_results is not None:
                    validation_results[nid] = report
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
