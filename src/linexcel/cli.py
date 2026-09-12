"""Command-line interface: ``linexcel analyze workbook.xlsx``.

Deterministic by default. AI documentation is opt-in via ``--ai-docs``, which
needs the optional ``openai`` client (``linexcel[ai]``) and an
OpenAI-compatible endpoint.
"""

from __future__ import annotations

import argparse
import contextlib
import sys
import threading
from collections.abc import Iterator
from pathlib import Path

from linexcel import __version__
from linexcel.i18n import LANGUAGES


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="linexcel",
        description="Data lineage analysis for Excel workbooks.",
    )
    parser.add_argument(
        "--version", action="version", version=f"linexcel {__version__}"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    analyze = sub.add_parser(
        "analyze",
        help="Analyze a workbook and write a standalone HTML viewer.",
        description=(
            "Analyze a workbook and write a standalone HTML viewer. "
            "Deterministic only unless --ai-docs is passed."
        ),
    )
    analyze.add_argument("workbook", type=Path, help="Path to an .xlsx/.xlsm file.")
    analyze.add_argument(
        "-o",
        "--output",
        type=Path,
        help="HTML output (default: <workbook>_lineage.html). '-' writes to stdout.",
    )
    analyze.add_argument(
        "--json",
        type=Path,
        dest="json_path",
        help="Also write the graph as JSON. '-' writes to stdout.",
    )
    analyze.add_argument(
        "--no-html",
        action="store_true",
        help="Skip the HTML viewer (use with --json).",
    )
    analyze.add_argument(
        "--language",
        default="en",
        choices=LANGUAGES,
        help="Viewer and AI prompt language (default: en).",
    )
    analyze.add_argument(
        "--refs-dir",
        type=Path,
        help=(
            "Folder holding the workbooks this one links to, and the add-ins "
            "whose VBA it calls. Without it, a cell reading another file is "
            "named but not resolved."
        ),
    )
    analyze.add_argument(
        "--screenshots",
        type=Path,
        metavar="DIR",
        help=(
            "Render each sheet to a PNG in DIR (needs LibreOffice) and show "
            "them in the report."
        ),
    )
    analyze.add_argument(
        "--time-budget",
        type=float,
        metavar="SECONDS",
        help=(
            "Ceiling on the step-by-step decomposition (default 300). Past it "
            "cells keep their values and lose only their breakdown, and the "
            "report says so. Raise it for a workbook of long formula chains; "
            "0 skips the decomposition entirely."
        ),
    )
    analyze.add_argument(
        "--target",
        action="append",
        metavar="SHEET!A1",
        help=(
            "Limit the analysis to the cells feeding this one (repeatable; "
            "several comma-separated cells accepted in one flag). The engine "
            "boots without a global evaluation, the upstream subgraph is "
            "traced and evaluated, and the rest of the workbook is omitted "
            "from the lineage. Without --target, everything is evaluated."
        ),
    )
    analyze.add_argument(
        "--rich",
        action="store_true",
        help=(
            "Print the end-of-run summary as a table (needs the 'progress' "
            "extra; plain text when rich is not installed). Affects stderr "
            "only — the HTML/JSON report is identical either way."
        ),
    )
    analyze.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Say what the file declares — sheets, size, linked workbooks, "
            "which ceilings will apply — and stop without analysing it."
        ),
    )
    analyze.add_argument(
        "-v", "--verbose", action="store_true", help="Per-phase timing on stderr."
    )

    mode = analyze.add_mutually_exclusive_group()
    mode.add_argument(
        "--deterministic-only",
        action="store_true",
        help="Lineage only, no AI calls (default).",
    )
    mode.add_argument(
        "--ai-docs",
        action="store_true",
        help=(
            "Also generate AI documentation. Requires the 'ai' extra: "
            "uvx --from 'linexcel[ai]' linexcel analyze --ai-docs ..."
        ),
    )

    ai = analyze.add_argument_group(
        "AI options (with --ai-docs or --vision-docs)",
        "Endpoint defaults come from LINEXCEL_AI_BASE_URL / LINEXCEL_AI_MODEL / "
        "LINEXCEL_AI_API_KEY. No provider is chosen implicitly.",
    )
    ai.add_argument("--base-url", help="OpenAI-compatible endpoint URL.")
    ai.add_argument("--model", help="Model name served by that endpoint.")
    ai.add_argument("--api-key", help="API key (prefer the env var).")
    ai.add_argument(
        "--max-workers", type=int, default=4, help="Concurrent requests (default: 4)."
    )
    ai.add_argument("--max-tokens", type=int, help="Cap output tokens per node.")
    ai.add_argument(
        "--token-budget",
        type=int,
        help="Cap total tokens (in + out) for the whole run.",
    )
    ai.add_argument(
        "--no-workbook-doc",
        action="store_true",
        help="Skip the workbook-level overview, document nodes only.",
    )
    ai.add_argument(
        "--vision-docs",
        action="store_true",
        help=(
            "Describe each sheet screenshot with a multimodal model. Needs "
            "--screenshots; independent of --ai-docs."
        ),
    )
    ai.add_argument(
        "--vision-model",
        help="Model that looks at the screenshots (default: --model).",
    )
    return parser


def _format_duration(seconds: float) -> str:
    """An order of magnitude, not a stopwatch reading.

    The estimate is derived from how much formula the file holds, on one
    machine; quoting it to the second would claim a precision it does not
    have, on hardware it knows nothing about.
    """
    if seconds < 90:
        return f"about {round(seconds / 5) * 5 or 5} seconds"
    return f"about {round(seconds / 60)} minute" + ("s" if seconds >= 90 else "")


#: A run lasting this many times its estimate is no longer within noise.
OVERRUN_FACTOR = 4
#: Never cry wolf sooner than this, however small the estimate was.
OVERRUN_NOTICE_FLOOR_SECONDS = 60.0


def _warn_if_long(workbook: Path) -> float:
    """Say how long this is likely to take, before it starts taking it.

    Returns the estimate so the caller can watch for the run blowing past it.
    Silence under a few seconds: a heads-up on every small file would be
    noise, and noise is what people learn to skip.
    """
    from linexcel.structure import (
        WORTH_MENTIONING_SECONDS,
        estimate_seconds,
        sheet_bytes,
    )

    try:
        data = workbook.read_bytes()
    except OSError:
        return 0.0  # the analysis itself will report this properly
    seconds = estimate_seconds(data)
    if seconds < WORTH_MENTIONING_SECONDS:
        return seconds
    weight = sheet_bytes(data)
    print(
        f"{weight / 1_048_576:.0f} MB of formulas: this should take "
        f"{_format_duration(seconds)} — an estimate, not a promise. It "
        f"weighs the formulas and counts them, but not how deep their "
        f"dependency chains run, and a workbook of long chains takes "
        f"longer. -v shows progress; --time-budget caps the part that can "
        f"run away.",
        file=sys.stderr,
    )
    return seconds


@contextlib.contextmanager
def _overrun_notice(estimated_seconds: float) -> Iterator[None]:
    """Print one actionable note if the run sails past its estimate.

    A timer thread, because the phases that run long are blocking Rust calls
    no Python-side loop can report from. Fires once, after the longer of four
    times the estimate and a minute; cancelled silently when the run finishes
    in time, which is nearly all of them.
    """
    wait = max(estimated_seconds * OVERRUN_FACTOR, OVERRUN_NOTICE_FLOOR_SECONDS)

    def _fire() -> None:
        context = (
            f" — already {wait / estimated_seconds:.0f}× the "
            f"{_format_duration(estimated_seconds)} estimate"
            if estimated_seconds >= 1
            else ""
        )
        print(
            f"[linexcel] still running{context}. Long dependency chains or "
            f"references the engine cannot resolve are the usual causes, and "
            f"the warnings printed at the end will say what was left out. On "
            f"a rerun, -v shows which phase is slow and --time-budget "
            f"SECONDS caps the step-by-step decomposition.",
            file=sys.stderr,
        )

    timer = threading.Timer(wait, _fire)
    timer.daemon = True
    timer.start()
    try:
        yield
    finally:
        timer.cancel()


def _report_dry_run(workbook: Path, facts: dict) -> None:
    """Print what the file claims, and what that means for the run.

    Written to stdout: unlike the progress and the statistics, this *is* the
    output of the command, and someone will pipe it.
    """
    ceilings = facts["ceilings"]
    print(f"{workbook.name}  {facts['bytes'] / 1_048_576:.1f} MB")
    print(f"{len(facts['sheets'])} sheet(s), {facts['declaredCells']:,} cells declared")
    for sheet in facts["sheets"]:
        state = "" if sheet["state"] == "visible" else f" [{sheet['state']}]"
        flag = "  ← over the ceiling, will be cut short" if sheet["truncated"] else ""
        print(
            f"  {sheet['name']}{state}: {sheet['rows']:,} × {sheet['cols']:,} "
            f"= {sheet['cells']:,} cells{flag}"
        )
    if facts["externalWorkbooks"]:
        names = ", ".join(facts["externalWorkbooks"])
        print(f"reads {len(facts['externalWorkbooks'])} other workbook(s): {names}")
        print("  pass --refs-dir DIR to resolve them against a folder")
    if facts["densePathRefused"]:
        print(
            "a sheet declares more than this can read densely: values will be "
            "read the slow way, and some may be missing. If the sheet does not "
            "really hold that much, delete the empty rows below and columns "
            "right of the data and save."
        )
    print(
        f"ceilings: {ceilings['cellsPerSheet']:,} cells and "
        f"{ceilings['nodesPerSheet']:,} nodes per sheet"
    )
    seconds = facts["estimatedSeconds"]
    print(
        f"analysing it should take {_format_duration(seconds)}"
        if seconds >= 1
        else "analysing it should take a moment"
    )


def _print_summary(result, *, fancy: bool) -> None:
    """The end-of-run tally: one plain line by default, a table when asked.

    ``fancy`` is opt-in (``--rich``) and degrades to the plain line when rich
    is not installed. Either way it writes to stderr only, so the HTML/JSON
    report is byte-identical with or without it.
    """
    line = (
        f"Nodes: {len(result.nodes)}  Edges: {len(result.edges)}  "
        f"Sheets: {len(result.sheets)}"
    )
    if not fancy:
        print(line, file=sys.stderr)
        return
    from linexcel.progress import _rich_console

    console = _rich_console()
    if console is None:
        print(
            "hint: --rich needs the 'progress' extra "
            "(pip install 'linexcel[progress]'); the plain summary follows.",
            file=sys.stderr,
        )
        print(line, file=sys.stderr)
        return
    from collections import Counter

    from rich.table import Table

    kinds = Counter(str(n.get("kind", "unknown")) for n in result.nodes)
    table = Table(title="linexcel summary")
    table.add_column("Node kind")
    table.add_column("Count", justify="right")
    for kind, count in sorted(kinds.items()):
        table.add_row(kind, f"{count:,}")
    table.add_section()
    table.add_row("total", f"{len(result.nodes):,}")
    console.print(table)
    console.print(
        f"Edges: {len(result.edges):,}  Sheets: {len(result.sheets):,}",
        highlight=False,
    )


def _report_recommendations(
    result,
    *,
    screenshots_asked: bool,
    screenshots_failed: bool,
    refs_dir: Path | None,
) -> None:
    """What a rerun could do better, one line each, printed last.

    Same register as the dry-run notes: plain advice about flags, on stderr,
    after the warnings that carry the detail.
    """
    hints: list[str] = []
    if screenshots_failed:
        hints.append(
            "the report has no sheet screenshots — fix the renderer named "
            "above and rerun with --screenshots DIR to include them"
        )
    elif not screenshots_asked and len(result.sheets) > 1:
        hints.append(
            f"{len(result.sheets)} sheets would render as pictures — rerun "
            "with --screenshots DIR to see them in the report "
            "(needs LibreOffice)"
        )
    stats = result.graph.get("meta", {}).get("stats", {})
    unread = stats.get("externalWorkbooks", 0) - stats.get("externalWorkbooksRead", 0)
    if (
        unread > 0
        and refs_dir is None
        and not any("--refs-dir" in w for w in result.warnings)
    ):
        hints.append(
            "pass --refs-dir DIR pointing at the other workbook(s) to read "
            "them instead of trusting the cache Excel left behind"
        )
    for hint in hints:
        print(f"hint: {hint}", file=sys.stderr)


def _run_analyze(args: argparse.Namespace) -> int:
    from linexcel import analyze as analyze_workbook
    from linexcel.structure import inspect_workbook

    if args.dry_run:
        _report_dry_run(args.workbook, inspect_workbook(args.workbook.read_bytes()))
        return 0
    estimated = _warn_if_long(args.workbook)

    if args.vision_docs and args.screenshots is None:
        raise ValueError(
            "--vision-docs describes the sheet screenshots, so it needs "
            "--screenshots DIR to render them first."
        )
    if args.vision_docs and args.deterministic_only:
        raise ValueError(
            "--vision-docs sends the screenshots to a model, which "
            "--deterministic-only rules out."
        )

    with _overrun_notice(estimated):
        result = analyze_workbook(
            args.workbook,
            verbose=args.verbose,
            refs_dir=args.refs_dir,
            step_seconds=args.time_budget,
            targets=args.target,
        )

    screenshots = None
    screenshot_error: str | None = None
    if args.screenshots is not None:
        from linexcel.insights import WorkbookRenderError

        try:
            screenshots = result.save_screenshots(args.screenshots)
        except WorkbookRenderError as exc:
            # The analysis itself is done and worth keeping: the report is
            # written without screenshots, and the message says why and how
            # to enable the renderer.
            screenshot_error = str(exc)
            print(f"warning: {exc}", file=sys.stderr)
        else:
            print(f"Shots: {args.screenshots}", file=sys.stderr)

    screenshot_docs: dict[str, str] | None = None
    if args.vision_docs and screenshots:
        screenshot_docs = result.describe_screenshots(
            screenshots,
            api_key=args.api_key,
            model=args.vision_model or args.model,
            base_url=args.base_url,
            language=args.language,
            max_tokens=args.max_tokens,
            token_budget=args.token_budget,
        )

    docs: dict[str, str] | None = None
    workbook_doc: str | None = None
    if args.ai_docs:
        docs = result.document(
            api_key=args.api_key,
            model=args.model,
            base_url=args.base_url,
            language=args.language,
            max_workers=args.max_workers,
            max_tokens=args.max_tokens,
            token_budget=args.token_budget,
        )
        if not args.no_workbook_doc:
            workbook_doc = result.document_workbook(
                api_key=args.api_key,
                model=args.model,
                base_url=args.base_url,
                language=args.language,
                max_tokens=args.max_tokens,
                token_budget=args.token_budget,
            )

    if args.json_path:
        payload = result.to_json(indent=2)
        if str(args.json_path) == "-":
            sys.stdout.write(payload + "\n")
        else:
            args.json_path.write_text(payload, encoding="utf-8")
            print(f"JSON:  {args.json_path}", file=sys.stderr)

    if not args.no_html:
        if args.output is not None and str(args.output) == "-":
            sys.stdout.write(
                result.to_html(
                    docs=docs,
                    workbook_doc=workbook_doc,
                    screenshots=screenshots,
                    screenshot_docs=screenshot_docs,
                    language=args.language,
                )
            )
        else:
            out = args.output or args.workbook.with_name(
                f"{args.workbook.stem}_lineage.html"
            )
            result.save_html(
                out,
                docs=docs,
                workbook_doc=workbook_doc,
                screenshots=screenshots,
                screenshot_docs=screenshot_docs,
                language=args.language,
            )
            print(f"HTML:  {out}", file=sys.stderr)

    _print_summary(result, fancy=args.rich)
    if args.ai_docs or args.vision_docs:
        print(f"AI:    {result.token_usage}", file=sys.stderr)
    for warning in result.warnings:
        print(f"warning: {warning}", file=sys.stderr)
    _report_recommendations(
        result,
        screenshots_asked=args.screenshots is not None,
        screenshots_failed=screenshot_error is not None,
        refs_dir=args.refs_dir,
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    try:
        return _run_analyze(args)
    except KeyboardInterrupt:
        return 130
    except (ValueError, OSError, RuntimeError) as exc:
        print(f"linexcel: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
