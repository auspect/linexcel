"""Command-line interface: ``linexcel analyze workbook.xlsx``.

Deterministic by default. AI documentation is opt-in via ``--ai-docs``, which
needs the optional ``openai`` client (``linexcel[ai]``) and an
OpenAI-compatible endpoint.
"""

from __future__ import annotations

import argparse
import sys
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


def _run_analyze(args: argparse.Namespace) -> int:
    from linexcel import analyze as analyze_workbook

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

    result = analyze_workbook(
        args.workbook, verbose=args.verbose, refs_dir=args.refs_dir
    )

    screenshots = None
    if args.screenshots is not None:
        screenshots = result.save_screenshots(args.screenshots)
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

    print(
        f"Nodes: {len(result.nodes)}  Edges: {len(result.edges)}  "
        f"Sheets: {len(result.sheets)}",
        file=sys.stderr,
    )
    if args.ai_docs or args.vision_docs:
        print(f"AI:    {result.token_usage}", file=sys.stderr)
    for warning in result.warnings:
        print(f"warning: {warning}", file=sys.stderr)
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
