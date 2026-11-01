#!/usr/bin/env python3
"""Manual end-to-end validation of linexcel, against a local model.

Two workbooks, analysed then documented, rendered to PNG and written out as
standalone HTML reports to open by hand:

``sales``   a realistic 4-sheet B2B commercial and margin review — executive
            KPIs, regional plan vs. actuals, 40 enterprise orders, product
            catalog, hidden columns, charts, comments and defined names.
``stress``  a much larger English workbook built to break things: every Excel
            error value, unresolvable and external references, sheet names
            needing quotes, a circular pair, formulas past the truncation
            limit, and cell text aimed at the report's HTML escaping.

The default provider is a **local** OpenAI-compatible endpoint (Ollama), so a
full run costs nothing and sends nothing off the machine. Nothing is documented
unless a provider answers: there is no implicit cloud fallback.

Neither fixture can contain VBA: ``openpyxl`` preserves a ``vbaProject.bin`` but
cannot author one, so macros need a workbook Excel itself wrote.
``tests/fixtures/macros.xlsm`` is one, and the pytest suite reads it; point
``--file`` at a workbook of your own to exercise the path on a real model.

``--check-max-tokens`` answers a question the test suite cannot: an
OpenAI-compatible endpoint is free to ignore ``max_tokens``, and only the
endpoint you actually use can say whether yours does. It documents one node
under several ceilings and reports which of them cut the response.

    uv run --extra ai --extra screenshots validate_manual.py  # both, AI + vision
    uv run validate_manual.py --workbook stress      # the hostile one
    uv run validate_manual.py --workbook both
    uv run validate_manual.py --file macros.xlsm     # your own workbook (VBA)
    uv run validate_manual.py --model <tag>          # another local model
    uv run validate_manual.py --no-vision            # partial diagnostic, exit 1
    uv run validate_manual.py --no-ai                # partial diagnostic, exit 1
    uv run validate_manual.py --token-budget 50000 --max-nodes 8
    uv run validate_manual.py --check-max-tokens     # does the endpoint obey?
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field, replace
from datetime import UTC, datetime
from pathlib import Path
from urllib.parse import urlsplit

try:
    import validation_workbooks
except ImportError:  # pragma: no cover - openpyxl missing
    print("❌ openpyxl is not installed. Run 'uv run validate_manual.py'.")
    sys.exit(1)

# Import the working tree rather than any installed linexcel: this script
# validates the code being edited.
sys.path.insert(0, str(Path(__file__).parent / "src"))
import linexcel  # noqa: E402
from linexcel.aidoc import AiDocError, OpenAICompatProvider  # noqa: E402
from linexcel.insights import _sheet_names, empty_sheet_render_exemptions  # noqa: E402

#: Local OpenAI-compatible endpoint. Ollama serves one at this address.
DEFAULT_BASE_URL = "http://localhost:11434/v1"
#: Small local model, enough to exercise the prompts without a cloud key.
DEFAULT_MODEL = "qwen3.8"
#: Ceiling on the whole run. The sales workbook needs far less; the point is to
#: show the knob exists before anyone points it at a real workbook and a paid API.
DEFAULT_TOKEN_BUDGET = 200_000

SCREENSHOTS_ROOT = Path("validation_screenshots")


def code_fingerprints() -> dict[str, str]:
    """Identify the exact dirty working tree, not only the last commit."""
    root = Path(__file__).resolve().parent
    paths = [
        root / "validate_manual.py",
        root / "validation_workbooks.py",
        root / "uv.lock",
    ]
    paths.extend(
        path
        for path in (root / "src").rglob("*")
        if path.is_file() and path.suffix in {".py", ".html", ".md", ".json"}
    )
    return {
        path.relative_to(root).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(paths)
        if path.is_file()
    }


@dataclass(frozen=True)
class Case:
    """One workbook and everything the run needs to know about it."""

    name: str
    headline: str
    build: Callable[[], bytes]
    workbook: Path
    languages: tuple[str, ...]
    checks: Sequence[str] = field(default_factory=tuple)
    #: False for a workbook the user supplied: it is read, never rewritten, and
    #: already carries the results Excel stored.
    generated: bool = True
    root: Path = Path(".")
    source_receipt: dict | None = None

    def output(self, language: str) -> Path:
        stem = "validate_out" if self.name == "sales" else f"validate_out_{self.name}"
        return self.root / f"{stem}_{language}.html"

    @property
    def screenshots_dir(self) -> Path:
        return self.root / "screenshots" / self.name


CASES = {
    "sales": Case(
        name="sales",
        headline="B2B commercial and margin review",
        build=validation_workbooks.build_sales_workbook,
        workbook=Path("fy26_commercial_review.xlsx"),
        languages=("fr", "en"),
        checks=(
            "'Workbook overview' — does the AI describe the *file*, not the graph?",
            "  It should name the title in B2, the hidden baseline/cost columns",
            "  and the FP&A comment on C5, none of which any formula reveals.",
            "'Sheets' — every sheet shows its own rendered image and its first",
            "  cells, alongside comments, frozen panes and merged ranges.",
            "'Graph' — select a node, read its card against the decomposition.",
        ),
    ),
    "stress": Case(
        name="stress",
        headline="hostile English workbook",
        build=validation_workbooks.build_stress_workbook,
        workbook=Path("validation_stress.xlsx"),
        languages=("en",),
        checks=(
            "'Errors' nodes — every Excel error value should render as #DIV/0!,",
            "  #N/A and so on, never as a raw engine object.",
            "'Cross Refs' — external links, 3-D and structured references become",
            "  grey 'external reference' nodes rather than silent wrong edges.",
            "  The circular pair (B15/B16) must not hang or vanish.",
            "'Hostile Text' — no dialog may appear, and the payloads must read",
            "  back verbatim: </script>, __TITLE__, $&, the U+2028 line.",
            "Sheet tabs — O'Brien's Café must be spelled with ONE apostrophe",
            "  everywhere, and Hidden Config / Very Hidden Archive must appear.",
        ),
    ),
}


def _case_for_file(path: Path) -> Case:
    """A case wrapping a workbook of the user's own.

    The generated fixtures cannot carry macros — ``openpyxl`` preserves a
    ``vbaProject.bin`` but cannot author one — so this is the route for
    exercising VBA extraction, the call graph and the sheet read/write edges
    against a real file.
    """
    return Case(
        name=re.sub(r"[^A-Za-z0-9]+", "-", path.stem).strip("-").lower() or "file",
        headline=f"your workbook: {path.name}",
        build=path.read_bytes,
        workbook=path,
        languages=("en",),
        generated=False,
        checks=(
            "'Graph' — VBA procedures appear as orange hexagons; their call",
            "  edges are dotted and their sheet reads/writes are orange.",
            "Select a VBA node — the panel shows the extracted source.",
            "Anything that looks wrong here is worth a bug report: this is a",
            "  real workbook, not a fixture written to be analysable.",
        ),
    )


def corpus_files(manifest_path: Path) -> tuple[list[Case], list[dict]]:
    """Verify a locally acquired corpus before any model request or rendering."""
    manifest_path = manifest_path.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("version") != 1 or not isinstance(manifest.get("files"), list):
        raise ValueError("Corpus manifest must have version 1 and files[]")
    cases, addins = [], []
    names = set(CASES)
    for entry in manifest["files"]:
        if not isinstance(entry, dict) or entry.get("role") not in {
            "workbook",
            "addin",
        }:
            raise ValueError("Each corpus entry requires role workbook or addin")
        path = (manifest_path.parent / entry["path"]).resolve()
        expected = entry.get("sha256", "")
        if not re.fullmatch(r"[0-9a-f]{64}", expected):
            raise ValueError(f"Missing SHA-256 for corpus file {path.name}")
        if hashlib.sha256(path.read_bytes()).hexdigest() != expected:
            raise ValueError(f"Corpus source fingerprint mismatch: {path.name}")
        if not str(entry.get("source_url", "")).startswith("https://"):
            raise ValueError(f"Corpus source requires an HTTPS receipt: {path.name}")
        receipt = {**entry, "path": str(path), "manifest": str(manifest_path)}
        if entry["role"] == "addin":
            if path.suffix.lower() not in {".xlam", ".xla"}:
                raise ValueError("Add-in corpus entries must be .xlam or .xla")
            artifact_name = (
                "addin-" + re.sub(r"[^A-Za-z0-9]+", "-", path.stem).strip("-").lower()
            )
            artifact_name += "-" + path.suffix.lower().lstrip(".")
            if artifact_name in names:
                raise ValueError(
                    f"Duplicate corpus add-in artifact name: {artifact_name}"
                )
            names.add(artifact_name)
            receipt["artifact_name"] = artifact_name
            addins.append(receipt)
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Workbook corpus entries must be .xlsx or .xlsm")
        case = _case_for_file(path)
        name = "public-" + case.name
        if name in names:
            raise ValueError(f"Duplicate corpus case name: {name}")
        names.add(name)
        cases.append(
            replace(
                case,
                name=name,
                headline=f"public corpus: {path.name}",
                source_receipt=receipt,
            )
        )
    if not cases:
        raise ValueError("Corpus manifest contains no real-world workbook")
    return cases, addins


def recalculate(data: bytes, suffix: str = ".xlsx") -> bytes | None:
    """Round-trip through LibreOffice so the workbook carries stored results.

    ``openpyxl`` writes formulas but never their results, so a generated fixture
    has no cached values at all — and the report's "value stored in the file"
    against "value linexcel recalculated" comparison, the whole point of reading
    values back, has nothing to compare. Every real Excel file carries the last
    computed result; this makes the fixture behave like one.

    It also produces honest disagreements: LibreOffice does not implement every
    function formualizer does, so the file ends up storing ``#NAME?`` where
    linexcel computes a number. That is exactly the case the comparison exists
    to surface, and no hand-written fixture would think to include it.

    Returns ``None`` when LibreOffice is unavailable; the caller keeps the
    original bytes.
    """
    from linexcel.insights import find_libreoffice

    office = find_libreoffice()
    if not office:
        return None
    with tempfile.TemporaryDirectory(prefix="linexcel-recalc-") as temp:
        root = Path(temp)
        source = root / f"workbook{suffix}"
        source.write_bytes(data)
        out_dir = root / "out"
        from linexcel.insights import prepare_readonly_office_profile

        prepare_readonly_office_profile(root / "profile")
        try:
            subprocess.run(
                [
                    office,
                    f"-env:UserInstallation={(root / 'profile').as_uri()}",
                    "--headless",
                    "--norestore",
                    "--convert-to",
                    "xlsx",
                    "--outdir",
                    str(out_dir),
                    str(source),
                ],
                check=True,
                capture_output=True,
                text=True,
                timeout=180,
            )
        except (subprocess.SubprocessError, OSError):
            return None
        produced = sorted(out_dir.glob("*.xlsx"))
        return produced[0].read_bytes() if produced else None


# ──────────────────────────────────────────────
# Local provider preflight
# ──────────────────────────────────────────────


def provider_api_key(base_url: str) -> str:
    """Do not forward a gateway credential to a separate loopback runtime."""
    if urlsplit(base_url).hostname in {"localhost", "127.0.0.1", "::1"}:
        return "not-needed"
    return os.getenv("LINEXCEL_AI_API_KEY") or os.getenv("OPENAI_API_KEY") or ""


class AcceptanceProvider(OpenAICompatProvider):
    """Keep non-secret completion receipts alongside this manual run."""

    def __init__(self, *, response_receipts=None, response_journal=None, **kwargs):
        self.response_receipts = response_receipts
        self.response_journal = response_journal
        self._receipt_lock = threading.Lock()
        super().__init__(**kwargs)

    def _checked_response(self, response, prompt, text, *, require_stop=False):
        usage = None
        status = "failed"
        receipt = None
        if self.response_receipts is not None or self.response_journal is not None:
            choices = getattr(response, "choices", [])
            receipt = {
                "status": "pending",
                "requested_model": self._model,
                "resolved_model": getattr(response, "model", None),
                "response_id": getattr(response, "id", None),
                "finish_reason": (
                    getattr(choices[0], "finish_reason", None) if choices else None
                ),
                "streaming": require_stop,
                "prompt_text_sha256": hashlib.sha256(
                    prompt.encode("utf-8")
                ).hexdigest(),
                "content_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
                "content_characters": len(text),
                "raw_response": text,
                "usage": None,
            }
            # Preserve received content before the completion gate can reject it.
            # Provider bodies, headers and credentials are never copied here.
            with self._receipt_lock:
                if self.response_receipts is not None:
                    self.response_receipts.append(receipt)
        try:
            output, usage = super()._checked_response(
                response, prompt, text, require_stop=require_stop
            )
            status = "complete"
            return output, usage
        except AiDocError as exc:
            usage = exc.usage
            raise
        finally:
            if receipt is not None:
                with self._receipt_lock:
                    receipt["status"] = status
                    receipt["usage"] = vars(usage).copy() if usage is not None else None
                    if self.response_journal is not None:
                        # Persist every received completion before a whole
                        # language can fail or a long run can be interrupted.
                        with self.response_journal.open(
                            "a", encoding="utf-8"
                        ) as stream:
                            stream.write(json.dumps(receipt, ensure_ascii=False) + "\n")


def acceptance_provider(
    args: argparse.Namespace, *, vision: bool = False, response_receipts=None
):
    """Explicit provider configuration shared by all acceptance requests."""
    base = (args.vision_base_url or args.base_url) if vision else args.base_url
    model = (args.vision_model or args.model) if vision else args.model
    extra_body = {}
    if getattr(args, "reasoning_split", False):
        extra_body["reasoning_split"] = True
    if getattr(args, "api_thinking", None) is not None:
        extra_body["thinking"] = {"type": args.api_thinking}
    if getattr(args, "chat_template_thinking", None) is not None:
        extra_body["chat_template_kwargs"] = {
            "enable_thinking": args.chat_template_thinking == "enabled"
        }
    return AcceptanceProvider(
        response_receipts=response_receipts,
        response_journal=getattr(args, "_receipt_journal", None),
        base_url=base,
        model=model,
        api_key=provider_api_key(base),
        client_kwargs={"max_retries": getattr(args, "api_retries", 0)},
        streaming=getattr(args, "api_stream", False),
        request_kwargs={"extra_body": extra_body} if extra_body else None,
    )


class _NoModelRedirect(urllib.request.HTTPRedirectHandler):
    """Discovery must not forward bearer credentials to a redirected origin."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def probe_local_models(base_url: str, timeout: float = 15.0) -> list[str] | None:
    """Model ids served at ``base_url``, or ``None`` if nothing answers.

    Uses the OpenAI-compatible ``/models`` route, which Ollama, vLLM and LM
    Studio all expose, so the check does not assume a particular runtime.
    """
    url = base_url.rstrip("/") + "/models"
    key = provider_api_key(base_url)
    request = urllib.request.Request(
        url, headers={"Authorization": f"Bearer {key}"} if key else {}
    )
    try:
        opener = urllib.request.build_opener(_NoModelRedirect())
        with opener.open(request, timeout=timeout) as response:
            payload = json.load(response)
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return None
    return [entry.get("id", "") for entry in payload.get("data", [])]


def check_local_provider(base_url: str, model: str) -> bool:
    """Report whether ``model`` can answer at ``base_url``; explain if it cannot."""
    served = probe_local_models(base_url)
    if served is None:
        print(f"   ⚠️  No OpenAI-compatible endpoint answered at {base_url}")
        print("      Start one, e.g. 'ollama serve', or pass --base-url / --no-ai.")
        return False
    # Ollama implies the ':latest' tag when a name carries none, and reports the
    # tagged id back, so an exact match would reject the very name it accepts.
    if model in served or f"{model}:latest" in served:
        print(f"   ✅ {model} is served at {base_url}")
        return True
    print(f"   ⚠️  '{model}' is not served at {base_url}")
    if served:
        print(f"      Available here: {', '.join(sorted(served))}")
        print(f"      Pull it with 'ollama pull {model}', or pass --model <name>.")
    else:
        print("      That endpoint serves no model at all.")
    return False


# ──────────────────────────────────────────────
# Steps
# ──────────────────────────────────────────────


def report_structure(result: linexcel.LineageResult) -> None:
    print("\n1. 🧮 Deterministic analysis")
    print(f"   Sheets : {', '.join(result.sheets)}")
    stats = result.stats
    print(
        f"   Stats  : {stats['totalFormulas']} formulas · {stats['totalNodes']} nodes"
        f" · {stats['totalEdges']} edges · {stats['vbaProcs']} VBA"
    )
    kinds: dict[str, int] = {}
    for node in result.nodes:
        kinds[node["kind"]] = kinds.get(node["kind"], 0) + 1
    print(f"   Kinds  : {', '.join(f'{k}={v}' for k, v in sorted(kinds.items()))}")

    # References the analyser could not resolve are the whole point of the
    # stress workbook, so name them rather than leaving them as a count.
    opaque = sorted(n["label"] for n in result.nodes if n["kind"] == "opaque")
    if opaque:
        print(f"   Unresolved ({len(opaque)}), expected for external/dynamic refs:")
        for label in opaque:
            print(f"     · {label}")
    for warning in result.warnings:
        print(f"   ⚠️  {warning}")


def report_context(result: linexcel.LineageResult) -> None:
    """Show the context the AI overview is grounded in — not just the graph."""
    print("\n2. 📋 Workbook context (openpyxl only, Excel is never launched)")
    for sheet in result.workbook_context["sheets"]:
        details = []
        if sheet.get("freeze_panes"):
            details.append(f"freeze {sheet['freeze_panes']}")
        if sheet.get("hidden_columns"):
            details.append(f"hidden {', '.join(sheet['hidden_columns'])}")
        if sheet.get("merged_ranges"):
            details.append(f"merged {len(sheet['merged_ranges'])}")
        suffix = f" — {' · '.join(details)}" if details else ""
        print(f"   '{sheet['name']}'{suffix}")
        for comment in sheet.get("comments", []):
            text = comment["text"].strip().replace("\n", " ")
            print(f"     💬 {comment['cell']} ({comment['author']}): {text}")


def render_screenshots(
    result: linexcel.LineageResult, case: Case
) -> dict | list | None:
    """Render each sheet to a PNG, keyed by the sheet it shows."""
    print("\n3. 📸 Sheet screenshots (LibreOffice headless + pdftoppm)")
    target = case.screenshots_dir
    try:
        if target.exists() and any(target.iterdir()):
            raise ValueError(f"Refusing to overwrite existing screenshots: {target}")
        screenshots = result.save_screenshots(target, dpi=200)
    except Exception as exc:
        print(f"   ⚠️  Not rendered: {exc}")
        return None
    if isinstance(screenshots, dict):
        print(f"   ✅ {len(screenshots)} sheet(s) rendered in {target.as_posix()}/")
        for name in screenshots:
            print(f"     📄 {name}")
    else:
        # The renderer did not give one page per sheet, so nothing ties a page
        # to a sheet: they go to the viewer's own tab as a flat list.
        print(f"   ⚠️  {len(screenshots)} page(s), not one per sheet — shown flat")
    return screenshots


def document(
    result: linexcel.LineageResult,
    args: argparse.Namespace,
    language: str,
    *,
    validation_results: dict | None = None,
) -> tuple[dict[str, str] | None, str | None]:
    """Document the workbook and its nodes, within the token budget."""
    node_ids = selected_nodes(result.nodes, args.max_nodes, args.seed)
    if args.max_nodes is not None:
        calculation = [
            n["id"] for n in result.nodes if n["kind"] in ("cell", "group", "vba")
        ]
        print(
            f"   - {language}: documenting {len(node_ids)} of "
            f"{len(calculation)} nodes (--max-nodes)"
        )
    workbook_doc = None
    try:
        receipts = (
            validation_results.setdefault("request_receipts", [])
            if validation_results is not None
            else None
        )
        with acceptance_provider(args, response_receipts=receipts) as provider:
            return _document_with_provider(
                result,
                args,
                language,
                provider,
                node_ids,
                validation_results=validation_results,
            )
    except AiDocError as exc:
        print(f"   ❌ {exc}")
        return None, workbook_doc


def _document_with_provider(
    result, args, language, provider, node_ids, *, validation_results=None
):
    workbook_doc = None
    stage = "overview"
    try:
        print(f"   - {language}: workbook overview…")
        workbook_doc = result.document_workbook(
            language=language,
            token_budget=args.token_budget,
            max_tokens=args.max_tokens,
            validation_results=validation_results,
            provider=provider,
        )
        print(f"   - {language}: node cards…")
        stage = "nodes"
        docs = result.document(
            node_ids,
            language=language,
            max_workers=args.max_workers,
            token_budget=args.token_budget,
            max_tokens=args.max_tokens,
            validation_results=validation_results,
            provider=provider,
        )
    except AiDocError as exc:
        if validation_results is not None:
            validation_results.setdefault("provider_failures", []).append(
                ai_failure_evidence(exc, stage=stage, language=language)
            )
        print(f"   ❌ {exc}")
        return None, workbook_doc
    print(f"   - {language}: {len(docs)} card(s) written")
    return docs, workbook_doc


def ai_failure_evidence(error: AiDocError, *, stage: str, language: str) -> dict:
    """Record failure types without gateway bodies, headers or credentials."""
    chain = []
    cause: BaseException | None = error
    while cause is not None and len(chain) < 8:
        item: dict[str, object] = {"type": type(cause).__name__}
        for attribute in ("errno", "winerror", "status_code"):
            value = getattr(cause, attribute, None)
            if isinstance(value, int):
                item[attribute] = value
        chain.append(item)
        cause = cause.__cause__
    usage = getattr(error, "usage", None)
    return {
        "stage": stage,
        "language": language,
        "status": "failed",
        "cause_chain": chain,
        "failed_transport_attempt_usage": "not reported by SDK",
        "sdk_attempt_count": "not reported by SDK",
        "response_usage_available": usage is not None,
        "response_usage_estimated": usage.estimated if usage is not None else None,
        "response_usage_reported": usage is not None and not usage.estimated,
    }


def describe(
    result: linexcel.LineageResult,
    screenshots,
    args: argparse.Namespace,
    language: str,
    *,
    response_receipts=None,
    validation_results=None,
    vision_results=None,
    evidence_dir=None,
) -> dict[str, str] | None:
    """Have a multimodal model look at the rendered sheets.

    The only check that cannot be made from the graph: whether the endpoint
    actually reads the picture. A model that answers "the image is blank" is
    telling you something about itself, which is exactly what a manual
    validation run is for.
    """
    if not screenshots:
        print("   ⚠️  no screenshots to describe")
        return None
    model = args.vision_model or args.model
    print(f"   - {language}: describing screenshots with {model}…")
    try:
        with acceptance_provider(
            args, vision=True, response_receipts=response_receipts
        ) as provider:
            seen = result.describe_screenshots(
                screenshots,
                provider=provider,
                language=language,
                max_tokens=args.max_tokens,
                token_budget=args.token_budget,
                validation_results=vision_results,
                evidence_dir=evidence_dir,
            )
    except AiDocError as exc:
        if validation_results is not None:
            validation_results.setdefault("provider_failures", []).append(
                ai_failure_evidence(exc, stage="vision", language=language)
            )
        print(f"   ❌ {exc}")
        return None
    for name, text in seen.items():
        first = " ".join(text.split())[:110]
        print(f"     · {name}: {first}…")
    return seen


def link_documentation_receipts(checks: dict, expected: list[str]) -> dict:
    """Bind both editorial phases to distinct actual provider completions.

    Successful generation is coverage evidence only. It is not a semantic
    acceptance of the final Markdown or of the reviewer's own verdict.
    """
    receipts = checks.get("request_receipts", [])
    assigned: set[int] = set()
    records = []
    for target in expected:
        report = checks.get(target, {})
        review = report.get("specialized_review", {})
        stages = review.get("stages", [])
        sent = report.get("sent_dossier", {})
        blob = sent.get("json_text")
        dossier_valid = (
            isinstance(blob, str)
            and sent.get("json_text_sha256")
            == hashlib.sha256(blob.encode()).hexdigest()
        )
        if dossier_valid:
            try:
                dossier_valid = isinstance(json.loads(blob), dict)
            except (ValueError, TypeError):
                dossier_valid = False
        valid = (
            dossier_valid
            and review.get("enabled") is True
            and review.get("status") == "completed"
            and [stage.get("phase") for stage in stages] == ["draft", "review"]
        )
        for stage in stages:
            stage["receipt_link_complete"] = False
            system, user, raw = (
                stage.get("system_prompt"),
                stage.get("user_prompt"),
                stage.get("raw_response"),
            )
            if not all(isinstance(value, str) for value in (system, user, raw)):
                valid = False
                continue
            prompt_sha = hashlib.sha256((system + "\n\n" + user).encode()).hexdigest()
            content_sha = hashlib.sha256(raw.encode()).hexdigest()
            matches = [
                (index, receipt)
                for index, receipt in enumerate(receipts)
                if receipt.get("prompt_text_sha256") == prompt_sha
                and receipt.get("content_sha256") == content_sha
            ]
            if len(matches) != 1:
                valid = False
                continue
            index, receipt = matches[0]
            response_id = receipt.get("response_id")
            complete = (
                dossier_valid
                and index not in assigned
                and isinstance(response_id, str)
                and bool(response_id)
                and receipt.get("status") == "complete"
                and receipt.get("finish_reason") == "stop"
                and receipt.get("raw_response") == raw
                and stage.get("status") == "complete"
                and stage.get("system_prompt_sha256")
                == hashlib.sha256(system.encode()).hexdigest()
                and stage.get("user_prompt_sha256")
                == hashlib.sha256(user.encode()).hexdigest()
                and stage.get("prompt_text_sha256") == prompt_sha
                and stage.get("content_sha256") == content_sha
                and stage.get("usage") == receipt.get("usage")
                and blob in user
                and useful_text(raw)
            )
            stage["receipt_link_complete"] = complete
            stage["response_id"] = response_id
            stage["provider_receipt_index"] = index
            receipt["documentation_evidence"] = {
                "target": target,
                "phase": stage.get("phase"),
                "dossier_sha256": (report.get("sent_dossier") or {}).get(
                    "json_text_sha256"
                ),
            }
            assigned.add(index)
            valid = valid and complete
        if stages:
            valid = valid and stages[-1].get("raw_response") == report.get(
                "raw_response"
            )
        if len(stages) == 2:
            draft = stages[0].get("raw_response")
            review_user = stages[1].get("user_prompt")
            valid = valid and (
                isinstance(draft, str)
                and bool(draft)
                and isinstance(review_user, str)
                and "Candidate to review (untrusted, not evidence):\n" + draft
                in review_user
                and "SPECIALIZED REVIEW PASS." in stages[1].get("system_prompt", "")
            )
        review["receipt_link_complete"] = valid
        records.append({"target": target, "stages": len(stages), "passed": valid})
    unassigned = [index for index in range(len(receipts)) if index not in assigned]
    return {
        "expected_documents": len(expected),
        "expected_stages": 2 * len(expected),
        "linked_stages": len(assigned),
        "unassigned_receipt_indices": unassigned,
        "documents": records,
        "factual_certification": False,
        "passed": (
            bool(records) and all(r["passed"] for r in records) and not unassigned
        ),
    }


def link_vision_receipts(checks: dict, receipts: list) -> None:
    """Bind receipts to exact native crop batches, including invalid responses."""
    assigned: set[int] = set()
    for report in checks.values():
        for page in report.get("pages", []):
            for batch in page.get("batches", []):
                if batch.get("resolution") == "deterministic_pixels":
                    # This is a local pixel decision, never a provider response.
                    batch["receipt_link_complete"] = False
                    continue
                attempts = batch.get("attempts")
                targets = attempts if isinstance(attempts, list) else [batch]
                batch["receipt_link_complete"] = False
                for target in targets:
                    target["receipt_link_complete"] = False
                    matches = [
                        (index, receipt)
                        for index, receipt in enumerate(receipts)
                        if receipt.get("prompt_text_sha256")
                        == target.get("prompt_text_sha256")
                        and receipt.get("content_sha256")
                        == target.get("content_sha256")
                    ]
                    if len(matches) != 1:
                        continue
                    index, receipt = matches[0]
                    if index in assigned:
                        continue
                    assigned.add(index)
                    response_id = receipt.get("response_id")
                    raw = target.get("raw_response")
                    receipt["vision_evidence"] = {
                        "batch_id": batch["id"],
                        "attempt": target.get("number"),
                        "selected": target is batch or target is targets[-1],
                        "page": page["page"],
                        "original_sha256": batch["original_sha256"],
                        "json_prompt_sha256": target["json_prompt_sha256"],
                        "images": batch["images"],
                    }
                    target["response_id"] = response_id
                    target["provider_receipt_index"] = index
                    if isinstance(attempts, list):
                        target["provider_receipt"] = {
                            key: value
                            for key, value in receipt.items()
                            if key != "vision_evidence"
                        }
                    target["receipt_link_complete"] = (
                        receipt.get("status") == "complete"
                        and receipt.get("finish_reason") == "stop"
                        and receipt.get("content_sha256")
                        == target.get("content_sha256")
                        and isinstance(response_id, str)
                        and bool(response_id)
                        and (
                            not isinstance(attempts, list)
                            or (
                                isinstance(raw, str)
                                and receipt.get("raw_response") == raw
                                and hashlib.sha256(raw.encode("utf-8")).hexdigest()
                                == target.get("content_sha256")
                                and receipt.get("usage") == target.get("usage")
                            )
                        )
                    )
                if isinstance(attempts, list) and attempts:
                    batch["response_id"] = attempts[-1].get("response_id")
                    batch["provider_receipt_index"] = attempts[-1].get(
                        "provider_receipt_index"
                    )
                    batch["provider_receipt"] = attempts[-1].get("provider_receipt")
                    batch["receipt_link_complete"] = attempts[-1].get(
                        "receipt_link_complete", False
                    )


def vision_attempts_complete(batch: dict) -> bool:
    """Retries retain every received attempt and count each usage exactly once."""
    if "attempts" not in batch:
        return True
    attempts = batch["attempts"]
    if not isinstance(attempts, list) or not 1 <= len(attempts) <= 2:
        return False
    if [attempt.get("number") for attempt in attempts] != list(
        range(1, len(attempts) + 1)
    ) or batch.get("final_attempt") != len(attempts):
        return False
    hashes = [attempt.get("prompt_text_sha256") for attempt in attempts]
    if not all(hashes) or len(set(hashes)) != len(hashes):
        return False
    selected = attempts[-1]
    mirrored = (
        "status",
        "user_prompt",
        "json_prompt_sha256",
        "prompt_text_sha256",
        "system_prompt_sha256",
        "content_sha256",
        "raw_response",
        "usage",
        "provider_receipt",
    )
    if any(batch.get(key) != selected.get(key) for key in mirrored):
        return False
    if selected.get("status") != "complete":
        return False
    totals = {"input_tokens": 0, "output_tokens": 0, "requests": 0}
    estimated = False
    for index, attempt in enumerate(attempts):
        if not attempt.get("receipt_link_complete"):
            return False
        if index < len(attempts) - 1 and (
            attempt.get("status") != "invalid_response"
            or not attempt.get("contract_error")
        ):
            return False
        prompt = attempt.get("user_prompt")
        if (
            not isinstance(prompt, str)
            or hashlib.sha256(prompt.encode("utf-8")).hexdigest()
            != attempt.get("json_prompt_sha256")
            or attempt.get("system_prompt_sha256")
            != selected.get("system_prompt_sha256")
        ):
            return False
        usage = attempt.get("usage")
        if not isinstance(usage, dict):
            return False
        for key in totals:
            value = usage.get(key)
            if type(value) is not int or value < 0:
                return False
            totals[key] += value
        estimated = estimated or bool(usage.get("estimated"))
    total_usage = batch.get("usage_total")
    return bool(
        isinstance(total_usage, dict)
        and all(total_usage.get(key) == value for key, value in totals.items())
        and total_usage.get("estimated") is estimated
    )


def deterministic_vision_batch_complete(batch: dict, page: dict) -> bool:
    """Recompute exact no-text pixel proofs; no model receipt may be invented."""
    from linexcel.vision_evidence import VisionEvidenceError, prepare_image

    usage = batch.get("usage_total")
    images = batch.get("images")
    if (
        not isinstance(usage, dict)
        or not isinstance(images, list)
        or not images
        or not all(isinstance(image, dict) for image in images)
    ):
        return False
    if (
        batch.get("status") != "complete"
        or batch.get("provider_request_required") is not False
        or batch.get("attempts") != []
        or any(
            key in batch
            for key in (
                "raw_response",
                "provider_receipt",
                "response_id",
                "user_prompt",
                "content_sha256",
                "prompt_text_sha256",
                "final_attempt",
            )
        )
        or any(
            type(usage.get(key)) is not int or usage[key] != 0
            for key in ("input_tokens", "output_tokens", "requests")
        )
    ):
        return False
    regions = {r.get("id"): r for r in page.get("evidence", {}).get("regions", [])}
    expected = []
    for image in images:
        bbox = image.get("bbox")
        if (
            not isinstance(bbox, list)
            or len(bbox) != 4
            or any(type(value) is not int for value in bbox)
            or not (0 <= bbox[0] < bbox[2] and 0 <= bbox[1] < bbox[3])
        ):
            return False
        stored = regions.get(image.get("id"), {})
        try:
            payload = Path(stored["path"]).read_bytes()
            native = prepare_image(payload)
        except (KeyError, OSError, VisionEvidenceError):
            return False
        if len(native.regions) != 1:
            return False
        region = native.regions[0]
        if (
            region.text_visibility == "not_determined"
            or native.sha256 != stored.get("sha256")
            or native.width != image.get("width")
            or native.height != image.get("height")
            or native.width != bbox[2] - bbox[0]
            or native.height != bbox[3] - bbox[1]
            or image.get("pixel_evidence") != region.pixel_evidence()
            or stored.get("pixel_evidence") != region.pixel_evidence()
        ):
            return False
        expected.append(
            {
                "region_id": image["id"],
                "readability": "no_visible_content",
                "observations": [],
                "observation_origin": "deterministic_pixels",
                "pixel_evidence": region.pixel_evidence(),
                "original_bbox": image["bbox"],
                "crop_sha256": native.sha256,
            }
        )
    recorded = [
        observation
        for observation in page.get("observations", [])
        if observation.get("region_id") in batch.get("region_ids", [])
    ]
    return bool(expected and batch.get("observations") == expected == recorded)


def vision_region_coverage(expected: list[str], checks: dict) -> dict:
    """Complete image coverage requires all exact native regions and receipts."""
    completed = []
    for name, report in checks.items():
        pages = report.get("pages", [])
        if (
            report.get("status") != "complete"
            or not pages
            or len(pages) != report.get("expected_pages")
        ):
            continue
        valid = True
        for page in pages:
            evidence = page.get("evidence", {})
            regions = evidence.get("regions", [])
            ids = [region.get("id") for region in regions]
            batches = page.get("batches", [])
            processed = [
                identifier
                for batch in batches
                for identifier in batch.get("region_ids", [])
            ]
            valid = bool(
                valid
                and evidence.get("pixel_coverage_complete")
                and page.get("region_coverage_complete")
                and regions
                and batches
                and len(ids) == len(set(ids))
                and sorted(page.get("completed_regions", [])) == sorted(ids)
                and sorted(processed) == sorted(ids)
                and all(
                    batch.get("status") == "complete"
                    and (
                        deterministic_vision_batch_complete(batch, page)
                        if batch.get("resolution") == "deterministic_pixels"
                        else batch.get("receipt_link_complete")
                        and vision_attempts_complete(batch)
                    )
                    and batch.get("original_sha256") == evidence.get("original_sha256")
                    and [image.get("id") for image in batch.get("images", [])]
                    == batch.get("region_ids")
                    and all(
                        any(
                            image.get("id") == region.get("id")
                            and image.get("sha256") == region.get("sha256")
                            and image.get("bbox") == region.get("bbox")
                            for region in regions
                        )
                        for image in batch.get("images", [])
                    )
                    for batch in batches
                )
            )
            source = page.get("source_path")
            if source is not None:
                source_path = Path(source)
                valid = bool(
                    valid
                    and source_path.is_file()
                    and hashlib.sha256(source_path.read_bytes()).hexdigest()
                    == evidence.get("original_sha256")
                )
                if valid:
                    from linexcel.vision_evidence import (
                        VisionEvidenceError,
                        prepare_image,
                    )

                    try:
                        native = prepare_image(source_path.read_bytes())
                    except VisionEvidenceError:
                        valid = False
                    else:
                        # A white crop cannot stand in for pixels from a
                        # nonuniform original, even with a self-consistent SHA.
                        valid = all(
                            [r.get(key) for r in regions]
                            == [r.metadata().get(key) for r in native.regions]
                            for key in ("id", "bbox", "sha256", "width", "height")
                        )
            for region in regions:
                path = Path(region.get("path", ""))
                valid = bool(
                    valid
                    and path.is_file()
                    and hashlib.sha256(path.read_bytes()).hexdigest()
                    == region.get("sha256")
                )
        if valid:
            completed.append(name)
    return {
        "expected": expected,
        "completed": sorted(completed),
        "missing": sorted(set(expected) - set(completed)),
        "unexpected": sorted(set(completed) - set(expected)),
        "passed": set(completed) == set(expected),
        "factual_certification": False,
    }


def selected_nodes(nodes: list[dict], limit: int | None, seed: int) -> list[str]:
    """Reproducible sampling across sheets instead of a prefix of one sheet."""
    groups: dict[str, list[str]] = {}
    for node in nodes:
        if node["kind"] in ("cell", "group", "vba"):
            groups.setdefault(node.get("sheet") or "", []).append(node["id"])
    rng = random.Random(seed)
    buckets = [sorted(groups[key]) for key in sorted(groups)]
    for bucket in buckets:
        rng.shuffle(bucket)
    selected = []
    while any(buckets):
        for bucket in buckets:
            if bucket:
                selected.append(bucket.pop())
    return selected if limit is None else selected[:limit]


def useful_text(value: object) -> bool:
    return (
        isinstance(value, str)
        and bool(value.strip())
        and "AI returned empty response" not in value
    )


def coverage(expected: list[str], actual: dict | None) -> dict:
    actual = actual or {}
    valid = {key for key, value in actual.items() if useful_text(value)}
    return {
        "expected": expected,
        "completed": sorted(valid),
        "missing": sorted(set(expected) - valid),
        "unexpected": sorted(valid - set(expected)),
        "passed": valid == set(expected),
    }


def screenshot_evidence(screenshots: dict | list | None) -> list[dict]:
    if isinstance(screenshots, dict):
        named = [
            (name, path)
            for name, paths in screenshots.items()
            for path in ([paths] if isinstance(paths, str | Path) else paths)
        ]
    else:
        named = [(Path(path).stem, path) for path in screenshots or []]
    evidence = []
    for name, source in named:
        from PIL import Image

        path = Path(source)
        payload = path.read_bytes()
        valid = False
        width = height = 0
        error = None
        try:
            with Image.open(path) as image:
                width, height = image.size
                valid = image.format == "PNG"
                image.verify()
            with Image.open(path) as image:
                image.load()
        except (OSError, ValueError, SyntaxError) as exc:
            valid = False
            error = str(exc)
        evidence.append(
            {
                "name": name,
                "path": str(path.resolve()),
                "sha256": hashlib.sha256(payload).hexdigest(),
                "width": width,
                "height": height,
                "passed": valid and width > 0 and height > 0,
                "error": error,
            }
        )
    return evidence


#: Ceilings the max_tokens check documents one node under. ``None`` is the
#: reference length. 500 is the everyday case, which a small local model rarely
#: reaches — it shows the argument breaks nothing, and nothing more. 48 is the
#: one that has to bite, and is the only one that can prove enforcement.
MAX_TOKENS_PROBES = (None, 500, 48)
#: Providers count a response's tokens their own way, and a runtime may overrun
#: a ceiling by a token or two. Past this the ceiling is being ignored, not
#: rounded.
MAX_TOKENS_SLACK = 8


def check_max_tokens(
    result: linexcel.LineageResult, args: argparse.Namespace, language: str
) -> None:
    """Document one node under several ceilings and report what came back.

    ``max_tokens`` is passed straight through to the endpoint, so whether it is
    honoured is a property of the runtime, not of linexcel — an OpenAI-compatible
    server is free to ignore it. That is precisely why this is worth running
    against the endpoint you actually use, and why it reports rather than
    asserts: the useful outcome is knowing, not passing.
    """
    node = next(
        (n for n in result.nodes if n.get("kind") in ("cell", "group", "vba")), None
    )
    if node is None:
        return
    print(f"\n6. 🎚  max_tokens — one node ({node['label']}) under three ceilings")

    measured: list[tuple[int | None, int, int]] = []
    for ceiling in MAX_TOKENS_PROBES:
        before = result.token_usage.output_tokens
        try:
            with acceptance_provider(args) as provider:
                cards = result.document(
                    [node["id"]],
                    language=language,
                    provider=provider,
                    max_tokens=ceiling,
                    token_budget=args.token_budget,
                )
        except AiDocError as exc:
            print(f"   ❌ {exc}")
            return
        produced = result.token_usage.output_tokens - before
        measured.append((ceiling, produced, len(cards.get(node["id"], ""))))

    counted = "estimated" if result.token_usage.estimated else "reported"
    print(f"   ceiling   output tokens ({counted})   characters")
    for ceiling, produced, chars in measured:
        label = "none" if ceiling is None else str(ceiling)
        print(f"   {label:>7}   {produced:>21}   {chars:>10}")

    # Response length alone can disprove enforcement when it overshoots the
    # ceiling, but it cannot prove enforcement. A near-ceiling result is only
    # suggestive unless the provider also reports a length/limit finish reason.
    suggestive = exceeded = False
    for ceiling, produced, _chars in measured:
        if ceiling is None:
            continue
        if produced > ceiling + MAX_TOKENS_SLACK:
            exceeded = True
            print(f"   ❌ {ceiling} exceeded ({produced} tokens) — not enforced")
        elif produced >= ceiling - MAX_TOKENS_SLACK:
            suggestive = True
            print(
                f"   ➖ {ceiling} stopped near the ceiling ({produced} tokens) — "
                "suggestive, but unproven without a finish reason"
            )
        else:
            print(
                f"   ➖ {ceiling} inconclusive — the model stopped at {produced} "
                "on its own, so nothing was cut"
            )
    if exceeded:
        print("   → this endpoint does not honour max_tokens.")
    elif suggestive:
        print(
            "   ⚠️  Some runs stopped near the ceiling, but without a provider "
            "finish reason that remains suggestive only."
        )
    else:
        print(
            "   ⚠️  No ceiling hit, so nothing is proven either way. Lower the "
            "smallest value in MAX_TOKENS_PROBES and run again."
        )


def run_case(
    case: Case, args: argparse.Namespace, use_ai: bool, use_vision: bool
) -> dict:
    """Export evidence and report coverage independently for every language."""
    started = time.monotonic()
    outcome = {"case": case.name, "languages": {}, "errors": []}
    print(f"\n═══ {case.name}: {case.headline} ═══")
    data = case.build()
    if case.source_receipt is not None:
        if hashlib.sha256(data).hexdigest() != case.source_receipt["sha256"]:
            raise ValueError("Corpus input changed after manifest verification")
        outcome["source_receipt"] = case.source_receipt
    if case.generated and not args.no_recalc:
        stored = recalculate(data)
        if stored is None:
            outcome["errors"].append("LibreOffice recalculation failed")
            print("   ⚠️  LibreOffice unavailable: the fixture keeps no stored")
            print("      values, so the file-vs-recalculated comparison is blank.")
        else:
            data = stored
    # A generated fixture is written out so it can be opened in Excel next to the
    # report; a workbook the user pointed us at is theirs, and is never rewritten.
    if case.generated:
        case.workbook.write_bytes(data)
    print(f"   Test workbook: {case.workbook.resolve()}")

    result = linexcel.analyze(data, filename=case.workbook.name)
    outcome["execution"] = result.graph.get("meta", {}).get("execution", {})
    outcome["workbook_sha256"] = hashlib.sha256(data).hexdigest()
    outcome["workbook"] = str(case.workbook.resolve())
    (case.root / f"{case.name}-graph.json").write_text(
        result.to_json(indent=2), encoding="utf-8"
    )
    if outcome["execution"].get("status", "completed") != "completed":
        outcome["errors"].append(
            "Analysis did not complete within its execution policy"
        )
        outcome["passed"] = False
        return outcome
    report_structure(result)
    report_context(result)
    skip_screenshots = getattr(args, "no_screenshots", False)
    screenshots = None if skip_screenshots else render_screenshots(result, case)
    outcome["screenshots_skipped"] = skip_screenshots
    outcome["screenshots"] = screenshot_evidence(screenshots)
    if skip_screenshots:
        outcome["errors"].append("Screenshots intentionally skipped (--no-screenshots)")
    elif not outcome["screenshots"] or not all(
        item["passed"] for item in outcome["screenshots"]
    ):
        outcome["errors"].append("No valid screenshots, or invalid PNG headers")
    image_names = [item["name"] for item in outcome["screenshots"]]
    expected_images = sorted(set(image_names))
    outcome["sheets_without_mapped_image"] = sorted(
        (set(_sheet_names(data)) | set(result.sheets)) - set(expected_images)
    )
    outcome["empty_sheet_exemptions"] = empty_sheet_render_exemptions(data)
    missing_images = set(outcome["sheets_without_mapped_image"]) - set(
        outcome["empty_sheet_exemptions"]
    )
    if not skip_screenshots and not isinstance(screenshots, dict):
        outcome["errors"].append(
            "Page-to-sheet mapping is unavailable; complete sheet coverage is unproven"
        )
    elif not skip_screenshots and missing_images:
        outcome["errors"].append(
            "Missing sheet screenshots: " + ", ".join(sorted(missing_images))
        )
    outcome["image_mapping"] = "sheet" if isinstance(screenshots, dict) else "page"
    expected_nodes = selected_nodes(result.nodes, args.max_nodes, args.seed)

    print("\n4. 💾 Reports")
    described = False
    for language in case.languages:
        documentation_checks: dict = {}
        docs, workbook_doc = (
            document(result, args, language, validation_results=documentation_checks)
            if use_ai
            else (None, None)
        )
        vision_receipts = []
        vision_checks: dict = {}
        seen = (
            describe(
                result,
                screenshots,
                args,
                language,
                response_receipts=vision_receipts,
                validation_results=documentation_checks,
                vision_results=vision_checks,
                evidence_dir=case.root / "vision_evidence" / case.name,
            )
            if use_vision
            else None
        )
        described = described or seen is not None
        documentation_phases = link_documentation_receipts(
            documentation_checks, ["workbook", *expected_nodes]
        )
        link_vision_receipts(vision_checks, vision_receipts)
        path = case.output(language)
        result.save_html(
            path,
            docs=docs,
            workbook_doc=workbook_doc,
            screenshots=screenshots,
            screenshot_docs=seen,
            language=language,
        )
        status = {
            "nodes": coverage(expected_nodes, docs),
            "overview": useful_text(workbook_doc),
            "documentation_phases": documentation_phases,
            "vision": coverage(expected_images, seen),
            "vision_regions": vision_region_coverage(expected_images, vision_checks),
            "report": str(path.resolve()),
        }
        status["passed"] = bool(
            use_ai
            and use_vision
            and expected_images
            and status["nodes"]["passed"]
            and status["overview"]
            and status["documentation_phases"]["passed"]
            and status["vision"]["passed"]
            and status["vision_regions"]["passed"]
        )
        outcome["languages"][language] = status
        evidence = {
            "docs": docs,
            "workbook_doc": workbook_doc,
            "screenshot_docs": seen,
            "coverage": status,
            "documentation_checks": documentation_checks,
            "vision_request_receipts": vision_receipts,
            "vision_checks": vision_checks,
        }
        (case.root / f"{case.name}-{language}-ai.json").write_text(
            json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"   {'✅' if status['passed'] else '❌ INCOMPLETE'} {path.resolve()}")
        (case.root / f"{case.name}-validation.json").write_text(
            json.dumps(outcome, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    if use_ai:
        print(f"\n5. 🧾 Token usage: {result.token_usage}")
        if result.token_usage.estimated:
            print("   (approximated — this endpoint reported no usage block)")
        if args.check_max_tokens:
            check_max_tokens(result, args, case.languages[0])

    print(f"\n   Open {case.output(case.languages[0]).name} and check:")
    for line in case.checks:
        print(f"   · {line}" if not line.startswith("  ") else f"   {line}")
    if described:
        # The one claim in the report that nothing deterministic backs: read it
        # against the picture it sits under, not against the graph.
        print("   · 'Sheets' — each description is badged as read from the image.")
        print("     Check it against that image: a weak vision model describes a")
        print("     plausible spreadsheet rather than this one.")
    outcome["seconds"] = round(time.monotonic() - started, 3)
    if not case.generated:
        if (
            hashlib.sha256(case.workbook.read_bytes()).hexdigest()
            != outcome["workbook_sha256"]
        ):
            outcome["errors"].append("Source workbook changed during acceptance")
    outcome["token_usage"] = str(result.token_usage)
    outcome["passed"] = not outcome["errors"] and all(
        item["passed"] for item in outcome["languages"].values()
    )
    (case.root / f"{case.name}-validation.json").write_text(
        json.dumps(outcome, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return outcome


def run_addin(receipt: dict, root: Path) -> dict:
    """Read all VBA source; export a reviewable graph without executing an add-in.

    This is a separate static-source check, not node AI/worksheet coverage.
    The dashboard and a procedure sample still require independent visual review.
    """
    from linexcel.macros import analyze_macro_addin

    path = Path(receipt["path"])
    data = path.read_bytes()
    if hashlib.sha256(data).hexdigest() != receipt["sha256"]:
        raise ValueError("Add-in input fingerprint mismatch")
    source = linexcel.read_macros(data, filename=path.name)
    name = receipt["artifact_name"]
    (root / f"{name}-source.json").write_text(
        json.dumps(source, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    payload = analyze_macro_addin(data, path.name)
    result = linexcel.LineageResult(
        payload["graph"], None, source_data=data, filename=path.name
    )
    report = root / f"{name}.html"
    result.save_html(report, language="en")
    (root / f"{name}-graph.json").write_text(result.to_json(indent=2), encoding="utf-8")
    unchanged = hashlib.sha256(path.read_bytes()).hexdigest() == receipt["sha256"]
    return {
        "source_receipt": receipt,
        "scope": "static_vba_source_only",
        "modules": len(source["modules"]),
        "procedures": len(source["procedures"]),
        "execution": "not_executed",
        "warnings": source["warnings"],
        "source_unchanged": unchanged,
        "report": str(report.resolve()),
        "passed": unchanged
        and source["status"] == "read"
        and bool(source["procedures"]),
        "manual_review": (
            "pending: dashboard tabs, provenance and sampled source procedures; "
            "no worksheet/AI coverage claimed"
        ),
    }


def main() -> int:
    # A redirected stdout on Windows defaults to cp1252, which cannot encode the
    # emoji below — nor the workbook's own accented sheet names.
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if reconfigure is not None:
        reconfigure(encoding="utf-8", errors="replace")

    # Loading a credential file is explicit: an unrelated .env must never
    # silently switch the default local acceptance run to a hosted provider.
    env_parser = argparse.ArgumentParser(add_help=False)
    env_parser.add_argument("--env-file", type=Path)
    env_args, _ = env_parser.parse_known_args()
    if env_args.env_file is not None:
        if not env_args.env_file.is_file():
            env_parser.error("--env-file must name an existing file")
        try:
            from dotenv import load_dotenv
        except ImportError:
            env_parser.error("--env-file requires python-dotenv (uv sync --dev)")
        load_dotenv(env_args.env_file, override=False)
    parser = argparse.ArgumentParser(description=__doc__, parents=[env_parser])
    parser.add_argument(
        "--base-url",
        default=(
            os.getenv("LINEXCEL_AI_BASE_URL") or os.getenv("OPENAI_BASE_URL")
            if env_args.env_file is not None
            else None
        )
        or DEFAULT_BASE_URL,
        help=f"OpenAI-compatible endpoint (default: {DEFAULT_BASE_URL})",
    )
    parser.add_argument(
        "--model", default=DEFAULT_MODEL, help=f"model id (default: {DEFAULT_MODEL})"
    )
    parser.add_argument(
        "--vision",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="describe rendered sheets with a multimodal model (default: enabled)",
    )
    parser.add_argument(
        "--vision-model",
        default=None,
        help="model that looks at the screenshots (default: --model)",
    )
    parser.add_argument(
        "--vision-base-url",
        help="separate vision endpoint (default: --base-url)",
    )
    parser.add_argument(
        "--reasoning-split",
        action="store_true",
        help="request separate reasoning_content from providers supporting this option",
    )
    parser.add_argument(
        "--token-budget",
        type=int,
        default=DEFAULT_TOKEN_BUDGET,
        help=f"ceiling on total tokens (default: {DEFAULT_TOKEN_BUDGET:,})",
    )
    parser.add_argument("--max-workers", type=int, default=1)
    parser.add_argument("--seed", type=int, default=20260912)
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="new output directory; refuses to overwrite a previous run",
    )
    parser.add_argument(
        "--max-nodes",
        type=int,
        default=None,
        help=(
            "document a reproducible sample spread across sheets "
            "(partial diagnostic run)"
        ),
    )
    parser.add_argument(
        "--workbook",
        choices=(*CASES, "both"),
        default="both",
        help="which fixture to run (default: both)",
    )
    parser.add_argument(
        "--corpus-manifest",
        type=Path,
        help=(
            "verified local corpus manifest: add every public workbook to the "
            "full run and read all add-in source separately"
        ),
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=None,
        help=(
            "run the whole pipeline against a workbook of your own instead of a "
            "fixture. The only way to exercise VBA end to end: openpyxl can "
            "preserve a vbaProject.bin but not author one, so no generated "
            "fixture can contain macros — point this at a real .xlsm"
        ),
    )
    parser.add_argument(
        "--no-recalc",
        action="store_true",
        help=(
            "skip the LibreOffice round-trip that gives a generated fixture the "
            "stored results a real Excel file carries"
        ),
    )
    parser.add_argument(
        "--max-tokens",
        type=int,
        default=None,
        help="cap each individual AI response (approximate, provider-dependent)",
    )
    parser.add_argument(
        "--api-stream",
        action="store_true",
        help="receive complete AI responses as a stream, with terminal usage",
    )
    parser.add_argument(
        "--api-retries",
        type=int,
        default=0,
        help="explicit SDK retries for transient API failures (default: no retry)",
    )
    parser.add_argument(
        "--api-thinking",
        choices=("disabled", "adaptive"),
        default=None,
        help=(
            "explicit provider thinking.type for documentation and vision "
            "(default: omit; requires model support)"
        ),
    )
    parser.add_argument(
        "--chat-template-thinking",
        choices=("enabled", "disabled"),
        default=None,
        help=(
            "explicit llama.cpp chat_template_kwargs.enable_thinking for "
            "documentation and vision (default: omit; requires template support)"
        ),
    )
    parser.add_argument(
        "--check-max-tokens",
        action="store_true",
        help=(
            "document one node under several max_tokens ceilings and report "
            "whether the endpoint enforces them (costs a few extra requests)"
        ),
    )
    parser.add_argument(
        "--no-ai", action="store_true", help="skip AI documentation entirely"
    )
    parser.add_argument(
        "--no-screenshots",
        action="store_true",
        help="skip all sheet rendering (partial diagnostic run)",
    )
    args = parser.parse_args()
    if args.api_retries < 0:
        parser.error("--api-retries must be non-negative")
    if args.max_nodes is not None and args.max_nodes < 1:
        parser.error("--max-nodes must be positive")
    # Refused rather than quietly dropped, as ``linexcel analyze`` refuses
    # --vision-docs with --deterministic-only: a screenshot is the largest
    # thing this script can send, and --no-ai is the promise that it sends
    # nothing at all.
    if args.vision and args.no_ai and "--vision" in sys.argv:
        parser.error(
            "--vision sends the screenshots to a model, which --no-ai rules out."
        )

    if args.no_screenshots and "--vision" in sys.argv:
        parser.error("--vision requires screenshots; remove --no-screenshots.")

    print("--- 📊 linexcel manual validation ---")
    if args.file is not None:
        args.file = args.file.resolve()
        if not args.file.is_file():
            print(f"❌ No such workbook: {args.file}")
            return 1
        cases = [_case_for_file(args.file)]
    elif args.workbook == "both":
        cases = list(CASES.values())
    else:
        cases = [CASES[args.workbook]]
    addins = []
    if args.corpus_manifest is not None:
        try:
            public_cases, addins = corpus_files(args.corpus_manifest)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            parser.error(f"Invalid corpus manifest: {exc}")
        cases.extend(public_cases)

    root = (
        args.output_dir
        or SCREENSHOTS_ROOT / "manual" / datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    ).resolve()
    try:
        root.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        parser.error(
            f"Output directory already exists: {root}. Choose a new directory."
        )
    cases = [
        replace(
            case,
            root=root,
            workbook=root / case.workbook.name if case.generated else case.workbook,
        )
        for case in cases
    ]
    manifest = {
        "started_utc": datetime.now(UTC).isoformat(),
        "configuration": vars(args).copy(),
        "code_sha256": code_fingerprints(),
        "python": sys.version,
        "cases": [],
        "addins": [],
        "status": "running",
        "manual_review": "pending: inspect tabs, sampled nodes, and image claims",
        "provider_receipt_journal": "provider_receipts.jsonl",
    }
    args._receipt_journal = root / "provider_receipts.jsonl"

    def save_manifest():
        (root / "validation.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    save_manifest()

    use_ai = not args.no_ai
    use_vision = args.vision and use_ai and not args.no_screenshots
    if not use_ai:
        print("   AI skipped (--no-ai): reports keep every deterministic tab.")
    elif not check_local_provider(args.base_url, args.model):
        print("   Continuing without AI.")
        use_ai = use_vision = False
    else:
        print(f"   Budget: {args.token_budget:,} tokens per workbook")
        # The model that looks at the images is the one to probe when it is not
        # the one that writes: a request nothing serves fails per language, per
        # workbook, after the screenshots have already been rendered.
        vision_base = args.vision_base_url or args.base_url
        vision_model = args.vision_model or args.model
        if use_vision and (vision_base != args.base_url or vision_model != args.model):
            use_vision = check_local_provider(vision_base, vision_model)
            if not use_vision:
                print("   Continuing without the screenshot descriptions.")

    for case in cases:
        try:
            outcome = run_case(case, args, use_ai, use_vision)
        except Exception as exc:
            outcome = {
                "case": case.name,
                "passed": False,
                "errors": [f"{type(exc).__name__}: {exc}"],
            }
            print(f"❌ {case.name}: {exc}")
        manifest["cases"].append(outcome)
        save_manifest()
    for receipt in addins:
        try:
            manifest["addins"].append(run_addin(receipt, root))
        except Exception as exc:
            manifest["addins"].append(
                {
                    "source_receipt": receipt,
                    "passed": False,
                    "errors": [f"{type(exc).__name__}: {exc}"],
                }
            )
        save_manifest()
    partial = (
        args.no_ai
        or not args.vision
        or args.max_nodes is not None
        or args.no_recalc
        or args.no_screenshots
    )
    passed = (
        all(item["passed"] for item in manifest["cases"])
        and all(item["passed"] for item in manifest["addins"])
        and not partial
    )
    manifest["status"] = "passed" if passed else "incomplete" if partial else "failed"
    manifest["finished_utc"] = datetime.now(UTC).isoformat()
    save_manifest()
    print(f"\nValidation {manifest['status']}: {root / 'validation.json'}")
    print(
        "Generation coverage does not certify factual correctness; "
        "manual review remains required."
    )
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
