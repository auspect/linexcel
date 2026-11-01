#!/usr/bin/env python3
"""Capture the README screenshots from a generated linexcel report.

The images in ``imgs/`` illustrate the viewer, so they go stale the moment
``viewer.html``, ``viewer.py`` or ``i18n.py`` changes and nobody notices — a
README showing an interface the package no longer has. This script captures
them from a real report, and records which sources they were captured from so
the ``readme-shots`` prek hook can tell when they no longer match.

Each shot drives the interface the way a reader would — clicking tabs, searching
for a node — rather than photographing the landing state four times:

    uv run python scripts/capture_viewer.py            # recapture imgs/
    uv run python scripts/capture_viewer.py --check    # are they still current?

Capturing needs Playwright (``uv pip install playwright && playwright install
chromium``); ``--check`` needs nothing and is what CI and the commit hook run.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
IMAGES_DIR = ROOT / "imgs"
MANIFEST = IMAGES_DIR / "manifest.json"
DEFAULT_REPORT = ROOT / "validation_screenshots" / "readme" / "validate_out_en.html"

#: Changing any of these changes what the report looks like.
WATCHED_SOURCES = (
    Path("scripts/capture_viewer.py"),
    Path("validation_workbooks.py"),
    Path("src/linexcel/viewer.py"),
    # The report itself. It left viewer.py for its own file, and it is what
    # actually decides how a screenshot looks, so watching only the Python
    # that loads it would watch the wrong thing.
    Path("src/linexcel/assets/viewer.html"),
    Path("src/linexcel/i18n.py"),
)

VIEWPORT_WIDTH, VIEWPORT_HEIGHT = 1440, 900
#: Retina-density PNGs: the README is read on high-DPI displays and GitHub
#: downscales, so capturing at 1x looks soft.
SCALE = 2
#: fcose has `animate: false`, but the layout still runs after boot.
LAYOUT_SETTLE_MS = 2500


@dataclass(frozen=True)
class Shot:
    """One README image and the interaction that sets the report up for it."""

    name: str
    caption: str
    #: Element ids to click, in order, before capturing.
    clicks: tuple[str, ...] = ()
    #: Text typed into the search box and submitted with Enter.
    search: str | None = None
    #: Capture this element rather than the viewport.
    selector: str | None = None
    #: Select a named source sheet, then optionally bring its render into view.
    sheet: str | None = None
    show_render: bool = False


SHOTS = (
    Shot(
        name="viewer_graph.png",
        caption="Dependency graph, workbook-wide",
        clicks=("lin-tab-graph", "lin-fit"),
    ),
    Shot(
        # Summary!C10 (Net Commercial Contribution) combines Gross Profit, Net
        # Revenue, RebateRate and Commission in a pure numeric
        # formula whose cached value agrees identically between LibreOffice and
        # Formualizer, while showcasing multi-step subexpression evaluation.
        # Values and provenance align cleanly without locale comma ambiguity.
        name="viewer_node_documented.png",
        caption="Selected KPI: source formula, values, provenance and calculation steps",
        clicks=("lin-tab-graph",),
        search="Summary!C10",
    ),
    Shot(
        name="viewer_workbook_overview.png",
        caption="Workbook summary: rendered Summary sheet",
        clicks=("lin-tab-sheets",),
        sheet="Summary",
        show_render=True,
    ),
    Shot(
        name="viewer_sheets_context.png",
        caption="Sales sheet context: source rendering, metadata and comments",
        clicks=("lin-tab-sheets",),
        sheet="Sales",
    ),
)


def hash_sources() -> dict[str, str]:
    """SHA-256 of every watched source, keyed by POSIX-style repo path."""
    digests = {}
    for relative in WATCHED_SOURCES:
        path = ROOT / relative
        # Git can check out identical sources with CRLF on Windows and LF on
        # Linux. Line-ending conversion must not make screenshots look stale.
        source = path.read_bytes().replace(b"\r\n", b"\n")
        digests[relative.as_posix()] = hashlib.sha256(source).hexdigest()
    return digests


def check() -> int:
    """Exit non-zero when the recorded sources no longer match the tree."""
    if not MANIFEST.exists():
        print(f"❌ {MANIFEST.relative_to(ROOT)} is missing — the README images have")
        print("   never been captured from a known revision of the viewer.")
        print("   Run: uv run python scripts/capture_viewer.py")
        return 1

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    recorded = manifest.get("sources", {})
    current = hash_sources()
    stale = [name for name, digest in current.items() if recorded.get(name) != digest]
    missing = [shot.name for shot in SHOTS if not (IMAGES_DIR / shot.name).exists()]
    changed = [
        shot.name
        for shot in SHOTS
        if (IMAGES_DIR / shot.name).exists()
        and manifest.get("image_sha256", {}).get(shot.name)
        != hashlib.sha256((IMAGES_DIR / shot.name).read_bytes()).hexdigest()
    ]

    if not stale and not missing and not changed:
        print("✅ README screenshots match the current viewer.")
        return 0
    if stale:
        print("❌ The viewer changed since the README screenshots were captured:")
        for name in stale:
            print(f"     {name}")
    if missing:
        print("❌ Missing screenshots: " + ", ".join(missing))
    if changed:
        print("❌ Unverified or changed screenshots: " + ", ".join(changed))
    print("\n   Regenerate them, then commit imgs/ with your change:")
    print(
        "     uv run --extra screenshots validate_manual.py --workbook both --no-ai --output-dir validation_screenshots/readme-<new-run>"
    )
    print(
        "     uv run --extra screenshots scripts/capture_viewer.py --report validation_screenshots/readme-<new-run>/validate_out_en.html"
    )
    print("\n   (If the change cannot affect the report, re-record the hashes with")
    print("    --accept instead.)")
    return 1


def accept() -> int:
    """Re-record the source hashes without recapturing."""
    if not MANIFEST.exists():
        print("❌ Capture the screenshots before accepting source-only changes.")
        return 1
    payload = json.loads(MANIFEST.read_text(encoding="utf-8"))
    for shot in SHOTS:
        path = IMAGES_DIR / shot.name
        if (
            not path.exists()
            or payload.get("image_sha256", {}).get(shot.name)
            != hashlib.sha256(path.read_bytes()).hexdigest()
        ):
            print(
                "❌ Capture missing, changed or unverified screenshots before accepting."
            )
            return 1
    payload["sources"] = hash_sources()
    payload["accepted_without_recapture"] = date.today().isoformat()
    MANIFEST.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"✅ {MANIFEST.relative_to(ROOT)} updated without recapturing.")
    return 0


def write_manifest(report: Path, content: dict | None = None) -> None:
    report = report.resolve()
    payload = {
        "captured": date.today().isoformat(),
        "sources": hash_sources(),
        "images": {shot.name: shot.caption for shot in SHOTS},
        "image_sha256": {
            shot.name: hashlib.sha256((IMAGES_DIR / shot.name).read_bytes()).hexdigest()
            for shot in SHOTS
        },
        "report": {
            "path": report.relative_to(ROOT).as_posix()
            if report.is_relative_to(ROOT)
            else report.name,
            "sha256": hashlib.sha256(report.read_bytes()).hexdigest(),
        },
        "viewport": {
            "width": VIEWPORT_WIDTH,
            "height": VIEWPORT_HEIGHT,
            "scale": SCALE,
        },
        "content": content or {},
    }
    MANIFEST.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )


def report_content(report: Path) -> dict[str, bool]:
    """Record AI presence from the exported graph without assuming a model run."""
    _before, separator, after = report.read_text(encoding="utf-8").partition(
        "  var GRAPH = "
    )
    if not separator:
        raise ValueError("Report does not contain the exported viewer graph")
    graph, _end = json.JSONDecoder().raw_decode(after)
    meta = graph.get("meta", {})
    return {
        "aiNodeDocumentation": any(node.get("doc") for node in graph["nodes"]),
        "aiWorkbookOverview": bool(meta.get("workbookDoc")),
        "aiImageDescriptions": any(meta.get("screenshotDocs", {}).values()),
    }


def capture(report: Path) -> int:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Playwright is not installed. Install it with:")
        print("     uv pip install playwright && uv run playwright install chromium")
        return 1

    if not report.exists():
        print(f"❌ No report at {report}. Generate one first:")
        print(
            "     uv run --extra screenshots validate_manual.py --workbook both --no-ai --output-dir validation_screenshots/readme"
        )
        print(
            "   Optional AI content is captured only when the supplied report contains it."
        )
        return 1

    IMAGES_DIR.mkdir(exist_ok=True)
    content = report_content(report)
    print(f"📸 Capturing {report.name} → {IMAGES_DIR.relative_to(ROOT)}/")

    with (
        tempfile.TemporaryDirectory(
            prefix=".viewer-capture-", dir=IMAGES_DIR
        ) as temporary,
        sync_playwright() as playwright,
    ):
        browser = playwright.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        # Written inline rather than hoisted to a constant: `viewport` wants
        # Playwright's ViewportSize TypedDict, which a module-level dict widens
        # away from.
        page = browser.new_page(
            viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
            device_scale_factor=SCALE,
        )
        page.goto(report.resolve().as_uri())
        page.wait_for_selector("#lin-cy canvas", timeout=30_000)
        page.wait_for_timeout(LAYOUT_SETTLE_MS)
        for shot in SHOTS:
            target = Path(temporary) / shot.name
            if not apply_shot(page, shot):
                print(f"   ❌ {shot.name} — the required tab or node is unavailable")
                browser.close()
                return 1
            element = page.query_selector(shot.selector) if shot.selector else None
            if element is not None:
                element.screenshot(path=str(target))
            else:
                page.screenshot(path=str(target))
            print(f"   ✅ {shot.name} — {shot.caption}")

        browser.close()
        for shot in SHOTS:
            # TemporaryDirectory has a private Windows ACL. Moving its files
            # preserves that ACL and can make committed images unreadable to
            # other users/tools. Write bytes into imgs/ so its ACL is inherited.
            (IMAGES_DIR / shot.name).write_bytes(
                (Path(temporary) / shot.name).read_bytes()
            )

    write_manifest(report, content)
    print(f"   ✅ {MANIFEST.relative_to(ROOT).as_posix()}")
    return 0


def apply_shot(page, shot: Shot) -> bool:
    """Drive the report into the state ``shot`` documents.

    Returns ``False`` when a required tab or node is unavailable. The Visual
    and Sheets tabs contain source evidence without requiring an AI provider.
    """
    for element_id in shot.clicks:
        button = page.query_selector(f"#{element_id}")
        if button is None or button.get_attribute("hidden") is not None:
            return False
        button.click()
        page.wait_for_timeout(400)

    if shot.search:
        search = page.query_selector("#lin-search")
        if search is None:
            return False
        search.fill(shot.search)
        search.press("Enter")
        # The current viewer presents selectable results. Complete that user
        # action before capturing the node panel, so the list cannot cover it.
        # Older reports select directly and have no results popup.
        results = page.query_selector("#lin-search-results")
        if results is not None and results.is_visible():
            match = results.query_selector("#lin-result-list button")
            if match is None:
                return False
            match.click()
        page.wait_for_timeout(1200)

    if shot.sheet:
        button = page.locator("#lin-sheets-sidebar").get_by_role(
            "button", name=f"📄 {shot.sheet}", exact=True
        )
        if button.count() != 1:
            return False
        button.click()
        preview = page.locator("#lin-sheet-details .lin-shot-viewer img").first
        if preview.count() != 1:
            return False
        preview.scroll_into_view_if_needed()
        page.wait_for_function(
            "() => { const image = document.querySelector('#lin-sheet-details .lin-shot-viewer img'); return image && image.complete && image.naturalWidth > 0; }"
        )
        if shot.show_render:
            page.locator("#lin-sheet-details .lin-shot-viewer").first.evaluate(
                "element => element.scrollIntoView({block: 'start'})"
            )
        else:
            page.locator(".lin-sheet-body").evaluate("element => element.scrollTop = 0")
        page.wait_for_timeout(200)

    return True


def main() -> int:
    # This runs from a commit hook and from CI, where stdout is a pipe and
    # defaults to cp1252 on Windows — which cannot encode the symbols below.
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if reconfigure is not None:
        reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--report",
        type=Path,
        default=DEFAULT_REPORT,
        help=f"report to capture (default: {DEFAULT_REPORT.relative_to(ROOT)})",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="verify the images match the current viewer, capture nothing",
    )
    parser.add_argument(
        "--accept",
        action="store_true",
        help="re-record the source hashes without recapturing",
    )
    args = parser.parse_args()

    if args.check:
        return check()
    if args.accept:
        return accept()
    return capture(args.report)


if __name__ == "__main__":
    sys.exit(main())
