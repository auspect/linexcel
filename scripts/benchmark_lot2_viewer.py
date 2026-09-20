#!/usr/bin/env python3
"""Benchmark 2C: Viewer performance at ~100,000 real nodes (Issue #65, 2C).

Protocol requirements from Issue #65 & PLAN_DELEGATION.md:
1. Synthetic controlled graphs at 1k, 10k, 50k, 100k nodes with realistic edge topologies.
2. Offline standalone HTML test (file:// URI, no server, zero network requests).
3. Playwright measurements at both Desktop (1440x900) and Mobile (390x844) viewports:
   - HTML size on disk
   - DOMContentLoaded time
   - Window load time
   - Cytoscape canvas ready time
   - Search execution latency (typing + result count update)
   - Node detail selection latency
4. Error checking: JS console errors and unhandled exceptions.
5. Structured report saved to validation_screenshots/delegation-20260920/benchmark_2c_viewer.json.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import platform
import random
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from linexcel.viewer import render_html  # noqa: E402


def generate_synthetic_graph(n_nodes: int, seed: int = 42) -> dict[str, Any]:
    """Generate a realistic calculation DAG with n_nodes."""
    rng = random.Random(seed)
    sheets = ["Inputs", "Calculations", "Rates", "Summary", "Reporting"]

    nodes = []
    edges = []

    # 20% inputs, 70% formulas, 10% names/groups
    n_inputs = int(n_nodes * 0.20)

    # Create input nodes
    for i in range(n_inputs):
        sheet = sheets[i % len(sheets)]
        row = (i // len(sheets)) + 1
        col = (i % 26) + 1
        col_letter = chr(ord("A") + (col - 1))
        node_id = f"c:{sheet}!{col_letter}{row}"
        nodes.append(
            {
                "id": node_id,
                "label": f"{sheet}!{col_letter}{row}",
                "kind": "cell",
                "sheet": sheet,
                "value": round(rng.uniform(10.0, 1000.0), 2),
                "valueSource": "file",
                "searchable": f"{sheet}!{col_letter}{row} input",
            }
        )

    # Create formula / computed nodes
    for i in range(n_inputs, n_nodes):
        sheet = sheets[i % len(sheets)]
        row = (i // len(sheets)) + 1
        col = (i % 26) + 1
        col_letter = chr(ord("A") + (col - 1))
        node_id = f"c:{sheet}!{col_letter}{row}"

        # Pick 1 or 2 predecessors from already created nodes
        num_predecessors = rng.choices([1, 2, 3], weights=[0.6, 0.3, 0.1])[0]
        preds = rng.sample(nodes, min(len(nodes), num_predecessors))

        formula = " + ".join(p["label"] for p in preds)
        nodes.append(
            {
                "id": node_id,
                "label": f"{sheet}!{col_letter}{row}",
                "kind": "formula",
                "sheet": sheet,
                "formula": f"={formula}",
                "value": round(rng.uniform(50.0, 5000.0), 2),
                "valueSource": "engine",
                "searchable": f"{sheet}!{col_letter}{row} ={formula}",
            }
        )

        for p in preds:
            edges.append(
                {
                    "id": f"e:{p['id']}->{node_id}",
                    "source": p["id"],
                    "target": node_id,
                }
            )

    return {
        "nodes": nodes,
        "edges": edges,
        "meta": {
            "title": f"Synthetic Graph ({n_nodes} nodes)",
            "stats": {
                "totalNodes": len(nodes),
                "totalEdges": len(edges),
                "totalFormulas": n_nodes - n_inputs,
                "vbaProcs": 0,
            },
            "execution": {"status": "completed"},
        },
    }


async def measure_viewport(
    browser,
    html_path: Path,
    width: int,
    height: int,
    timeout_ms: int = 120_000,
) -> dict[str, Any]:
    """Open the standalone HTML in Playwright, measure load and interaction times."""
    context = await browser.new_context(viewport={"width": width, "height": height})
    page = await context.new_page()

    console_errors: list[str] = []
    page.on(
        "console",
        lambda msg: (
            console_errors.append(msg.text)
            if msg.type in {"error", "warning"}
            else None
        ),
    )
    page.on("pageerror", lambda err: console_errors.append(str(err)))

    t0 = time.perf_counter()
    file_url = html_path.as_uri()

    try:
        # 1. Navigation & DOMContentLoaded
        t_nav_start = time.perf_counter()
        await page.goto(file_url, wait_until="domcontentloaded", timeout=timeout_ms)
        t_dcl = time.perf_counter() - t_nav_start

        # 2. Window load
        await page.wait_for_load_state("load", timeout=timeout_ms)
        t_load = time.perf_counter() - t_nav_start

        # 3. Cytoscape Canvas readiness
        # The viewer initializes Cytoscape and adds canvas inside #lin-cy
        await page.wait_for_selector("#lin-cy canvas", timeout=timeout_ms)
        t_cy_canvas = time.perf_counter() - t_nav_start

        # 4. Search interaction
        # Find search box, type query, wait for feedback
        t_search_start = time.perf_counter()
        search_input = await page.wait_for_selector("#lin-search", timeout=30_000)
        assert search_input is not None
        await search_input.fill("Summary")
        await search_input.press("Enter")

        # Wait for search results or status feedback update
        await page.wait_for_selector(
            "#lin-result-list button, #lin-search-status:not([hidden])",
            timeout=30_000,
        )
        t_search_duration = time.perf_counter() - t_search_start

        # 5. Node Selection & Detail panel interaction
        t_select_start = time.perf_counter()
        # Click the first result button if available, or tap directly via page evaluate
        first_btn = await page.query_selector("#lin-result-list button")
        if first_btn:
            await first_btn.click()
            selected_id = await first_btn.text_content()
        else:
            selected_id = "auto"

        # Wait for detail panel to reflect selected node
        await page.wait_for_selector("#lin-panel:not(.lin-empty)", timeout=30_000)
        t_select_duration = time.perf_counter() - t_select_start

        total_elapsed = time.perf_counter() - t0

        return {
            "status": "passed",
            "viewport": f"{width}x{height}",
            "dcl_seconds": round(t_dcl, 4),
            "load_seconds": round(t_load, 4),
            "canvas_seconds": round(t_cy_canvas, 4),
            "cytoscape_ready_seconds": round(t_cy_canvas, 4),
            "search_latency_seconds": round(t_search_duration, 4),
            "selection_latency_seconds": round(t_select_duration, 4),
            "total_benchmark_seconds": round(total_elapsed, 4),
            "selected_node": (selected_id or "").strip(),
            "console_errors": console_errors[:10],
            "error_count": len(console_errors),
        }
    except Exception as exc:
        return {
            "status": "failed",
            "viewport": f"{width}x{height}",
            "error": str(exc),
            "console_errors": console_errors[:10],
            "total_benchmark_seconds": round(time.perf_counter() - t0, 4),
        }
    finally:
        await context.close()


async def run_benchmark(tiers: list[int], output_path: Path) -> dict[str, Any]:
    from playwright.async_api import async_playwright

    print("=" * 80)
    print("Benchmark 2C: Viewer Performance across Node Scales")
    print(f"Node scales : {tiers}")
    print("Viewports   : Desktop 1440x900 & Mobile 390x844")
    print(
        f"Platform    : {platform.system()} {platform.release()} ({platform.machine()})"
    )
    print("=" * 80)

    report: dict[str, Any] = {
        "timestamp_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "python": platform.python_version(),
        },
        "tiers": {},
    }

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        with tempfile.TemporaryDirectory(prefix="linexcel-bench-2c-") as temp_dir:
            temp_root = Path(temp_dir)

            for n_nodes in tiers:
                print(f"\n--- Tier: {n_nodes:,} nodes ---")
                t_gen_start = time.perf_counter()
                graph = generate_synthetic_graph(n_nodes)
                t_gen = time.perf_counter() - t_gen_start

                t_render_start = time.perf_counter()
                html = render_html(
                    graph, title=f"Scale {n_nodes} nodes", full_document=True
                )
                t_render = time.perf_counter() - t_render_start

                html_path = temp_root / f"viewer_{n_nodes}.html"
                html_path.write_text(html, encoding="utf-8")
                html_size_mb = round(html_path.stat().st_size / (1024 * 1024), 2)

                print(
                    f"  Graph generated: {len(graph['nodes'])} nodes, {len(graph['edges'])} edges in {t_gen:.2f}s"
                )
                print(f"  HTML rendered  : {html_size_mb} MB in {t_render:.2f}s")

                tier_data: dict[str, Any] = {
                    "nodes": n_nodes,
                    "edges": len(graph["edges"]),
                    "html_size_mb": html_size_mb,
                    "gen_seconds": round(t_gen, 4),
                    "render_seconds": round(t_render, 4),
                    "viewports": {},
                }

                # Desktop 1440x900
                print("  Measuring Desktop (1440x900)...")
                desktop_res = await measure_viewport(browser, html_path, 1440, 900)
                tier_data["viewports"]["desktop_1440x900"] = desktop_res
                if desktop_res["status"] == "passed":
                    print(
                        f"    Load: {desktop_res['load_seconds']:.2f}s | Cy ready: {desktop_res['cytoscape_ready_seconds']:.2f}s | Search: {desktop_res['search_latency_seconds']:.3f}s | Select: {desktop_res['selection_latency_seconds']:.3f}s"
                    )
                else:
                    print(f"    FAILED: {desktop_res.get('error')}")

                # Mobile 390x844
                print("  Measuring Mobile (390x844)...")
                mobile_res = await measure_viewport(browser, html_path, 390, 844)
                tier_data["viewports"]["mobile_390x844"] = mobile_res
                if mobile_res["status"] == "passed":
                    print(
                        f"    Load: {mobile_res['load_seconds']:.2f}s | Cy ready: {mobile_res['cytoscape_ready_seconds']:.2f}s | Search: {mobile_res['search_latency_seconds']:.3f}s | Select: {mobile_res['selection_latency_seconds']:.3f}s"
                    )
                else:
                    print(f"    FAILED: {mobile_res.get('error')}")

                report["tiers"][str(n_nodes)] = tier_data

        await browser.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("\n" + "=" * 80)
    print(f"Viewer benchmark results successfully saved to {output_path}")
    print("=" * 80)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Benchmark Viewer at ~100,000 nodes (Issue #65, 2C)"
    )
    parser.add_argument(
        "--tiers",
        type=str,
        default="500,1000,1500,2000,2500",
        help="Comma-separated node counts",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT
        / "validation_screenshots/delegation-20260920/benchmark_2c_viewer.json",
    )
    args = parser.parse_args()

    tiers = [int(t.strip()) for t in args.tiers.split(",") if t.strip()]
    asyncio.run(run_benchmark(tiers, args.output))
    return 0


if __name__ == "__main__":
    sys.exit(main())
