#!/usr/bin/env python3
"""Capture the README screenshots from a generated linexcel report.

The images in ``imgs/`` illustrate the viewer, so they go stale the moment
``viewer.py`` or ``i18n.py`` changes and nobody notices — a README showing an
interface the package no longer has. This script captures them from a real
report, and records which sources they were captured from so the ``readme-shots``
prek hook can tell when they no longer match.

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
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
IMAGES_DIR = ROOT / "imgs"
MANIFEST = IMAGES_DIR / "manifest.json"
DEFAULT_REPORT = ROOT / "validate_out_en.html"

#: Changing any of these changes what the report looks like.
WATCHED_SOURCES = (
    Path("src/linexcel/viewer.py"),
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


SHOTS = (
    Shot(
        name="viewer_graph.png",
        caption="Dependency graph, workbook-wide",
        clicks=("lin-tab-graph", "lin-fit"),
    ),
    Shot(
        # C5 rather than C7: both are documented, but C7 concatenates a number
        # into text and LibreOffice writes the French decimal comma when it
        # recalculates the fixture, so its card shows a genuine — and genuinely
        # confusing — disagreement over `OK: 6,7k` against `OK: 6.7k`. C5 agrees
        # exactly, which is what the value card is meant to demonstrate here.
        name="viewer_node_documented.png",
        caption="A node selected: formula, step-by-step evaluation, AI card",
        clicks=("lin-tab-graph",),
        search="Synthese!C5",
    ),
    Shot(
        name="viewer_workbook_overview.png",
        caption="AI overview, grounded in the deterministic dossier",
        clicks=("lin-tab-overview",),
    ),
    Shot(
        name="viewer_sheets_context.png",
        caption="Sheet context: comments, frozen panes, rendered pages",
        clicks=("lin-tab-sheets",),
    ),
)


def hash_sources() -> dict[str, str]:
    """SHA-256 of every watched source, keyed by POSIX-style repo path."""
    digests = {}
    for relative in WATCHED_SOURCES:
        path = ROOT / relative
        digests[relative.as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return digests


def check() -> int:
    """Exit non-zero when the recorded sources no longer match the tree."""
    if not MANIFEST.exists():
        print(f"❌ {MANIFEST.relative_to(ROOT)} is missing — the README images have")
        print("   never been captured from a known revision of the viewer.")
        print("   Run: uv run python scripts/capture_viewer.py")
        return 1

    recorded = json.loads(MANIFEST.read_text(encoding="utf-8")).get("sources", {})
    current = hash_sources()
    stale = [name for name, digest in current.items() if recorded.get(name) != digest]
    missing = [shot.name for shot in SHOTS if not (IMAGES_DIR / shot.name).exists()]

    if not stale and not missing:
        print("✅ README screenshots match the current viewer.")
        return 0
    if stale:
        print("❌ The viewer changed since the README screenshots were captured:")
        for name in stale:
            print(f"     {name}")
    if missing:
        print("❌ Missing screenshots: " + ", ".join(missing))
    print("\n   Regenerate them, then commit imgs/ with your change:")
    print("     uv run python validate_manual.py --model <local-model>")
    print("     uv run python scripts/capture_viewer.py")
    print("\n   (If the change cannot affect the report, re-record the hashes with")
    print("    --accept instead.)")
    return 1


def accept() -> int:
    """Re-record the source hashes without recapturing."""
    write_manifest()
    print(f"✅ {MANIFEST.relative_to(ROOT)} updated without recapturing.")
    return 0


def write_manifest() -> None:
    payload = {
        "captured": date.today().isoformat(),
        "sources": hash_sources(),
        "images": {shot.name: shot.caption for shot in SHOTS},
    }
    MANIFEST.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )


def capture(report: Path) -> int:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Playwright is not installed. Install it with:")
        print("     uv pip install playwright && uv run playwright install chromium")
        return 1

    if not report.exists():
        print(f"❌ No report at {report}. Generate one first:")
        print("     uv run python validate_manual.py --model <local-model>")
        print("   The AI tabs stay empty without a provider, so the overview shot")
        print("   would show an interface the README describes as documented.")
        return 1

    IMAGES_DIR.mkdir(exist_ok=True)
    print(f"📸 Capturing {report.name} → {IMAGES_DIR.relative_to(ROOT)}/")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
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
            target = IMAGES_DIR / shot.name
            if not apply_shot(page, shot):
                print(f"   ⏭️  {shot.name} — its tab is absent from this report")
                continue
            element = page.query_selector(shot.selector) if shot.selector else None
            if element is not None:
                element.screenshot(path=str(target))
            else:
                page.screenshot(path=str(target))
            print(f"   ✅ {shot.name} — {shot.caption}")

        browser.close()

    write_manifest()
    print(f"   ✅ {MANIFEST.relative_to(ROOT).as_posix()}")
    return 0


def apply_shot(page, shot: Shot) -> bool:
    """Drive the report into the state ``shot`` documents.

    Returns ``False`` when the shot's tab does not exist in this report — an
    overview tab needs an AI run, and a half-rendered screenshot is worse than
    a missing one.
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
        page.wait_for_timeout(1200)

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
        help=f"report to capture (default: {DEFAULT_REPORT.name})",
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
