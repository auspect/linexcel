#!/usr/bin/env python3
"""Capture visual review screenshots of generated HTML reports.

Captures at Desktop (1440x900) and Mobile (390x844).
"""

import asyncio
from pathlib import Path

from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "validation_screenshots/manual-lot1"
OUT_DIR = BASE_DIR / "review_captures"
OUT_DIR.mkdir(parents=True, exist_ok=True)

REPORTS = [
    ("sales_en", BASE_DIR / "validate_out_en.html"),
    ("sales_fr", BASE_DIR / "validate_out_fr.html"),
    ("stress_en", BASE_DIR / "validate_out_stress_en.html"),
]

VIEWPORTS = [
    ("desktop", 1440, 900),
    ("mobile", 390, 844),
]


async def capture_all():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        for rep_name, rep_path in REPORTS:
            if not rep_path.exists():
                print(f"Skipping {rep_name}, {rep_path} does not exist")
                continue

            file_url = rep_path.as_uri()

            for vp_name, width, height in VIEWPORTS:
                print(f"Capturing {rep_name} at {vp_name} ({width}x{height})...")
                context = await browser.new_context(
                    viewport={"width": width, "height": height},
                    device_scale_factor=1,
                )
                page = await context.new_page()

                # 1. Graph tab (default landing or switch)
                await page.goto(file_url, wait_until="networkidle")
                await page.wait_for_timeout(2500)
                await page.screenshot(
                    path=str(OUT_DIR / f"{rep_name}_{vp_name}_graph.png"),
                    full_page=False,
                )

                # Select a node to populate detail panel
                try:
                    await page.evaluate("""() => {
                        if (window.cy && window.cy.nodes().length > 0) {
                            const n = window.cy.nodes()[0];
                            n.emit('tap');
                            return n.id();
                        }
                        return null;
                    }""")
                    await page.wait_for_timeout(1000)
                    await page.screenshot(
                        path=str(OUT_DIR / f"{rep_name}_{vp_name}_node_detail.png"),
                        full_page=False,
                    )
                except Exception as e:
                    print(f"Could not select node: {e}")

                # 2. Overview tab
                try:
                    await page.click("#lin-tab-overview")
                    await page.wait_for_timeout(1000)
                    await page.screenshot(
                        path=str(OUT_DIR / f"{rep_name}_{vp_name}_overview.png"),
                        full_page=False,
                    )
                except Exception as e:
                    print(f"Could not switch to overview: {e}")

                # 3. Sheets tab
                try:
                    await page.click("#lin-tab-sheets")
                    await page.wait_for_timeout(1000)
                    await page.screenshot(
                        path=str(OUT_DIR / f"{rep_name}_{vp_name}_sheets.png"),
                        full_page=False,
                    )
                except Exception as e:
                    print(f"Could not switch to sheets: {e}")

                await context.close()

        await browser.close()
    print("Captures completed in", OUT_DIR)


if __name__ == "__main__":
    asyncio.run(capture_all())
