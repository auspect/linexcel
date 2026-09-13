#!/usr/bin/env python3
"""Manual end-to-end validation of linexcel, against a local model.

Two workbooks, analysed then documented, rendered to PNG and written out as
standalone HTML reports to open by hand:

``sales``   a small readable sales report — a table that does not start at A1,
            a hidden column, merged cells, a chart, a comment, a defined name,
            cross-sheet aggregation. This is what the README screenshots show.
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
import random
import re
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field, replace
from datetime import UTC, datetime
from pathlib import Path

try:
    import validation_workbooks
except ImportError:  # pragma: no cover - openpyxl missing
    print("❌ openpyxl is not installed. Run 'uv run validate_manual.py'.")
    sys.exit(1)

# Import the working tree rather than any installed linexcel: this script
# validates the code being edited.
sys.path.insert(0, str(Path(__file__).parent / "src"))
import linexcel  # noqa: E402
from linexcel.aidoc import AiDocError  # noqa: E402
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

    def output(self, language: str) -> Path:
        stem = "validate_out" if self.name == "sales" else f"validate_out_{self.name}"
        return self.root / f"{stem}_{language}.html"

    @property
    def screenshots_dir(self) -> Path:
        return self.root / "screenshots" / self.name


CASES = {
    "sales": Case(
        name="sales",
        headline="readable sales report",
        build=validation_workbooks.build_sales_workbook,
        workbook=Path("validation_demo.xlsx"),
        languages=("fr", "en"),
        checks=(
            "'Workbook overview' — does the AI describe the *file*, not the graph?",
            "  It should name the title in B2, the hidden Price column and the",
            "  comment on B3, none of which any formula reveals.",
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


def probe_local_models(base_url: str, timeout: float = 3.0) -> list[str] | None:
    """Model ids served at ``base_url``, or ``None`` if nothing answers.

    Uses the OpenAI-compatible ``/models`` route, which Ollama, vLLM and LM
    Studio all expose, so the check does not assume a particular runtime.
    """
    url = base_url.rstrip("/") + "/models"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
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
    result: linexcel.LineageResult, args: argparse.Namespace, language: str
) -> tuple[dict[str, str] | None, str | None]:
    """Document the workbook and its nodes, within the token budget."""
    provider = {"base_url": args.base_url, "model": args.model}
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
        print(f"   - {language}: workbook overview…")
        workbook_doc = result.document_workbook(
            language=language,
            token_budget=args.token_budget,
            max_tokens=args.max_tokens,
            **provider,
        )
        print(f"   - {language}: node cards…")
        docs = result.document(
            node_ids,
            language=language,
            max_workers=args.max_workers,
            token_budget=args.token_budget,
            max_tokens=args.max_tokens,
            **provider,
        )
    except AiDocError as exc:
        print(f"   ❌ {exc}")
        return None, workbook_doc
    print(f"   - {language}: {len(docs)} card(s) written")
    return docs, workbook_doc


def describe(
    result: linexcel.LineageResult,
    screenshots,
    args: argparse.Namespace,
    language: str,
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
        seen = result.describe_screenshots(
            screenshots,
            base_url=args.base_url,
            model=model,
            language=language,
            max_tokens=args.max_tokens,
            token_budget=args.token_budget,
        )
    except AiDocError as exc:
        print(f"   ❌ {exc}")
        return None
    for name, text in seen.items():
        first = " ".join(text.split())[:110]
        print(f"     · {name}: {first}…")
    return seen


def selected_nodes(nodes: list[dict], limit: int | None, seed: int) -> list[str]:
    """Reproducible sampling across sheets instead of a prefix of one sheet."""
    groups: dict[str, list[str]] = {}
    for node in nodes:
        if node["kind"] in ("cell", "group", "vba"):
            groups.setdefault(node.get("sheet", ""), []).append(node["id"])
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
            cards = result.document(
                [node["id"]],
                language=language,
                base_url=args.base_url,
                model=args.model,
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
    outcome["workbook_sha256"] = hashlib.sha256(data).hexdigest()
    outcome["workbook"] = str(case.workbook.resolve())
    (case.root / f"{case.name}-graph.json").write_text(
        result.to_json(indent=2), encoding="utf-8"
    )
    report_structure(result)
    report_context(result)
    screenshots = render_screenshots(result, case)
    outcome["screenshots"] = screenshot_evidence(screenshots)
    if not outcome["screenshots"] or not all(
        item["passed"] for item in outcome["screenshots"]
    ):
        outcome["errors"].append("No valid screenshots, or invalid PNG headers")
    image_names = [item["name"] for item in outcome["screenshots"]]
    expected_images = sorted(set(image_names))
    if len(image_names) != len(expected_images):
        outcome["errors"].append(
            "Multiple pages for one sheet: "
            "vision currently describes only its first image"
        )
    outcome["sheets_without_mapped_image"] = sorted(
        (set(_sheet_names(data)) | set(result.sheets)) - set(expected_images)
    )
    outcome["empty_sheet_exemptions"] = empty_sheet_render_exemptions(data)
    missing_images = set(outcome["sheets_without_mapped_image"]) - set(
        outcome["empty_sheet_exemptions"]
    )
    if not isinstance(screenshots, dict):
        outcome["errors"].append(
            "Page-to-sheet mapping is unavailable; complete sheet coverage is unproven"
        )
    elif missing_images:
        outcome["errors"].append(
            "Missing sheet screenshots: " + ", ".join(sorted(missing_images))
        )
    outcome["image_mapping"] = "sheet" if isinstance(screenshots, dict) else "page"
    expected_nodes = selected_nodes(result.nodes, args.max_nodes, args.seed)

    print("\n4. 💾 Reports")
    described = False
    for language in case.languages:
        docs, workbook_doc = (
            document(result, args, language) if use_ai else (None, None)
        )
        seen = describe(result, screenshots, args, language) if use_vision else None
        described = described or seen is not None
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
            "vision": coverage(expected_images, seen),
            "report": str(path.resolve()),
        }
        status["passed"] = bool(
            use_ai
            and use_vision
            and expected_images
            and status["nodes"]["passed"]
            and status["overview"]
            and status["vision"]["passed"]
        )
        outcome["languages"][language] = status
        evidence = {
            "docs": docs,
            "workbook_doc": workbook_doc,
            "screenshot_docs": seen,
            "coverage": status,
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
    outcome["token_usage"] = str(result.token_usage)
    outcome["passed"] = not outcome["errors"] and all(
        item["passed"] for item in outcome["languages"].values()
    )
    (case.root / f"{case.name}-validation.json").write_text(
        json.dumps(outcome, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return outcome


def main() -> int:
    # A redirected stdout on Windows defaults to cp1252, which cannot encode the
    # emoji below — nor the workbook's own accented sheet names.
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if reconfigure is not None:
        reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-url",
        default=DEFAULT_BASE_URL,
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
    args = parser.parse_args()
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
        "configuration": vars(args),
        "code_sha256": code_fingerprints(),
        "python": sys.version,
        "cases": [],
        "status": "running",
        "manual_review": "pending: inspect tabs, sampled nodes, and image claims",
    }

    def save_manifest():
        (root / "validation.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    save_manifest()

    use_ai = not args.no_ai
    use_vision = args.vision and use_ai
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
        if use_vision and args.vision_model and args.vision_model != args.model:
            use_vision = check_local_provider(args.base_url, args.vision_model)
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
    partial = (
        args.no_ai or not args.vision or args.max_nodes is not None or args.no_recalc
    )
    passed = all(item["passed"] for item in manifest["cases"]) and not partial
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
