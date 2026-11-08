"""Capture the real local dashboard at three widths, with layout evidence.

Run a dedicated server on 8792 first. Artifacts, including imported projects,
remain in validation_screenshots. No AI requests are made by this probe.
"""

from __future__ import annotations

import argparse
import atexit
import hashlib
import json
import re
import subprocess
import time
from datetime import UTC, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--project")
    parser.add_argument(
        "--data-dir", type=Path, default=Path("validation_screenshots/frontend-server")
    )
    parser.add_argument(
        "--server-log",
        type=Path,
        default=Path("validation_screenshots/frontend-server.log"),
    )
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    output = args.output or Path("validation_screenshots/frontend") / datetime.now(
        UTC
    ).strftime("%Y%m%dT%H%M%S%fZ")
    output.mkdir(parents=True)
    log = subprocess.check_output(
        [
            "uv",
            "run",
            "--extra",
            "screenshots",
            "linexcel",
            "session-link",
            "--data-dir",
            str(args.data_dir),
        ]
    )
    log_text = log.decode("utf-16" if log.startswith(b"\xff\xfe") else "utf-8")
    url = re.search(r"http://127\.0\.0\.1:879[2-9]/#session=\w+", log_text)[0]
    manifest = {
        "status": "failed_incomplete",
        "workbook": str(args.workbook.resolve()),
        "screenshots": [],
        "errors": [],
        "metrics": [],
        "sources": {
            str(p): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in [
                Path("src/linexcel/assets") / name
                for name in ["app.html", "app-graph.js", "app-graph.css"]
            ]
        },
    }
    atexit.register(
        lambda: (output / "validation.json").write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    )
    with sync_playwright() as driver:
        browser = driver.chromium.launch()
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()
        page.on("pageerror", lambda error: manifest["errors"].append(str(error)))
        page.goto(url)
        page.wait_for_function(
            "() => document.querySelector('#connection').textContent.includes('connecté')"
        )

        def api(path, body=None):
            return page.evaluate(
                """async ([path,body])=>{const r=await fetch('api/'+path,body?{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)}:{});const data=await r.json();if(!r.ok)throw Error(JSON.stringify(data));return data}""",
                [path, body],
            )

        created = (
            {"project": {"id": args.project}}
            if args.project
            else api(
                "projects/local",
                {
                    "path": str(args.workbook.resolve()),
                    "name": "Frontend validation",
                    "budget": {"seconds": 600, "memoryMb": 1024},
                },
            )
        )
        project = created["project"]["id"]
        manifest["project"] = project
        deadline = time.monotonic() + 620
        while time.monotonic() < deadline:
            state = api("projects/" + project)
            if state["tasks"] and all(
                t["status"] not in ("queued", "running") for t in state["tasks"]
            ):
                break
            page.wait_for_timeout(300)
        manifest["importTasks"] = state["tasks"]
        page.evaluate("id=>localStorage.setItem('linexcel-project',id)", project)
        page.goto(url.split("#")[0] + "?lang=en")
        page.locator("#project-title").get_by_text(state["project"]["name"]).wait_for()
        page.evaluate(
            "() => {document.querySelector('#seconds').value=1800;document.querySelector('#import-seconds').value=1800;document.querySelector('#memory').value=2048}"
        )
        nodes = state["graph"]["nodes"]
        target = next(
            n["id"]
            for n in nodes
            if n.get("formula") and n.get("kind") in {"cell", "formula"}
        )
        manifest["selectedNode"] = target

        def capture(name, width, height):
            page.wait_for_timeout(250)
            path = output / f"{width}-{name}.png"
            page.screenshot(path=str(path), full_page=False)
            manifest["screenshots"].append(str(path.resolve()))
            metrics = page.evaluate(
                """()=>({overflow:document.documentElement.scrollWidth>innerWidth,width:innerWidth,canvas:(()=>{const e=document.querySelector('#graph-canvas');const b=e.getBoundingClientRect();return {width:b.width,height:b.height}})(), notices:[...document.querySelectorAll('.notice')].filter(e=>e.checkVisibility()&&e.textContent.trim()).map(e=>({id:e.id,text:e.textContent.slice(0,120)}))})"""
            )
            metrics["tabBar"] = page.locator("nav.tabs").bounding_box()
            if (
                name.startswith("graph")
                or name.startswith("neighborhood")
                or name == "focus"
            ):
                metrics["geometry"] = (
                    page.evaluate("""()=>{const stage=document.querySelector('#graph-canvas'),box=stage.getBoundingClientRect(),flow=stage._flow;
                const bounds=e=>{const b=e.getBoundingClientRect();return {x:b.x,y:b.y,width:b.width,height:b.height}};
                const nodes=[...stage.querySelectorAll('.react-flow__node')].map(e=>({id:e.dataset.id,...bounds(e)}));
                const boxes=flow?.getNodes().map(n=>flow.getNodeBounds(n.id))||[];const crossings=[...stage.querySelectorAll('.react-flow__edge-path')].filter(path=>{const length=path.getTotalLength();for(let i=1;i<200;i++){const p=path.getPointAtLength(length*i/200);if(boxes.some(b=>p.x>b.x1+2&&p.x<b.x2-2&&p.y>b.y1+2&&p.y<b.y2-2))return true;}return false;}).length;return {crossings,viewport:flow?.getViewport(),nodes,clipped:nodes.filter(n=>n.x<box.x-1||n.y<box.y-1||n.x+n.width>box.right+1||n.y+n.height>box.bottom+1).map(n=>n.id),panels:[...stage.querySelectorAll('.react-flow__panel')].filter(e=>e.checkVisibility()).map(e=>({class:e.className,...bounds(e)})),reader:document.querySelector('#graph-preview').checkVisibility()?bounds(document.querySelector('#graph-preview')):null,visibleSteps:[...document.querySelectorAll('#graph-calculation-steps .step-card')].filter(e=>{const b=e.getBoundingClientRect(),p=document.querySelector('#graph-panel-calculation').getBoundingClientRect();return e.checkVisibility()&&b.top>=p.top&&b.bottom<=p.bottom}).length}}""")
                )
            if name.startswith("neighborhood") and "group-cells" not in name:
                assert metrics["geometry"]["crossings"] == 0, (
                    name,
                    width,
                    metrics["geometry"],
                )
                if width == 1440:
                    assert metrics["geometry"]["viewport"]["zoom"] >= 0.45, (
                        name,
                        metrics["geometry"],
                    )
            if name == "auto-decomposition":
                metrics["firstStep"] = page.locator(
                    "#calculation-steps .step-card"
                ).first.bounding_box()
                metrics["inspector"] = page.locator("#node-detail").bounding_box()
            if name == "graph":
                metrics["frames"] = page.evaluate("""() => new Promise(resolve=>{
                  const intervals=[];let previous=performance.now();
                  function frame(now){intervals.push(now-previous);previous=now;
                    if(intervals.length<60)requestAnimationFrame(frame);
                    else{const sorted=intervals.slice(1).sort((a,b)=>a-b);
                      resolve({maxMs:Math.max(...sorted),p95Ms:sorted[Math.floor(sorted.length*.95)]});}}
                  requestAnimationFrame(frame);
                })""")
            manifest["metrics"].append(
                {"name": name, "viewport": [width, height], **metrics}
            )

        for width, height in [(1440, 900), (2560, 1440), (390, 844)]:
            page.set_viewport_size({"width": width, "height": height})
            page.goto(url.split("#")[0] + "?lang=en")
            page.locator("#project-title").get_by_text(
                state["project"]["name"]
            ).wait_for()
            page.evaluate(
                "() => {document.querySelector('#seconds').value=1800;document.querySelector('#memory').value=2048}"
            )
            page.wait_for_timeout(300)
            for tab in [
                "nodes",
                "sheets",
                "data",
                "captures",
                "tasks",
                "diagnostics",
                "graph",
            ]:
                page.locator("#tab-" + tab).click()
                if tab == "graph":
                    page.wait_for_function(
                        "() => document.querySelectorAll('#graph-breadcrumbs button').length>0",
                        timeout=620_000,
                    )
                    if page.locator("#graph-breadcrumbs button").count() > 1:
                        page.locator("#graph-breadcrumbs button").first.click()
                    page.locator(".react-flow__node").first.wait_for()
                if tab == "data":
                    page.wait_for_timeout(700)
                capture(tab, width, height)
            # Choose a real sheet with copied formulas, avoiding the one-node
            # defined-name group that previously hid leaf-layout defects.
            root = api(f"projects/{project}/graph-view")
            candidates = []
            for node in root["nodes"]:
                if node.get("level") != "sheet":
                    continue
                view = api(f"projects/{project}/graph-view?view={node['viewId']}")
                patterns = [n for n in view["nodes"] if n.get("level") == "pattern"]
                if patterns:
                    candidates.append((max(n["nodeCount"] for n in patterns), node))
            sheet = max(candidates, key=lambda item: item[0])[1]
            page.locator("#graph-canvas").evaluate(
                "(e,id)=>e._flow.select(id)", sheet["id"]
            )
            page.wait_for_timeout(300)
            capture("graph-patterns", width, height)
            for depth in range(12):
                cards = page.locator("#graph-canvas").evaluate(
                    "(e)=>e._flow.getNodes().map(n=>n.data.original)"
                )
                groups = [
                    n
                    for n in cards
                    if n.get("viewId") is not None and n.get("kind") != "outside"
                ]
                group = max(
                    (n for n in groups if n.get("level") == "pattern"),
                    key=lambda n: n["nodeCount"],
                    default=None,
                ) or max(groups, key=lambda n: n["nodeCount"], default=None)
                if group is None:
                    break
                page.locator("#graph-canvas").evaluate(
                    "(e,id)=>e._flow.select(id)", group["id"]
                )
                page.wait_for_timeout(300)
            capture("graph-cells", width, height)
            page.locator("#tab-nodes").click()
            if page.locator("#search").is_hidden():
                page.locator(".node-return").click()
            page.locator("#search").fill(target)
            page.locator(f'[data-node="{target}"]').first.click()
            page.wait_for_function(
                "() => !document.querySelector('#recalculate').disabled && document.querySelector('.step-card')",
                timeout=1_830_000,
            )
            capture("auto-decomposition", width, height)
            page.locator("#tab-graph").click()
            page.locator("#graph-neighborhood").click()
            page.locator("#graph-depth").evaluate(
                "e=>{e.value=1;e.dispatchEvent(new Event('change'))}"
            )
            page.wait_for_function(
                "()=>document.querySelector('#graph-canvas')._flow?.getNodes().some(n=>n.id===document.querySelector('#graph-search').value)||document.querySelector('#graph-preview').checkVisibility()"
            )
            page.wait_for_timeout(500)
            capture("neighborhood-reader", width, height)
            if page.locator("#graph-depth").is_hidden():
                page.locator("#graph-options>summary").click()
            page.locator("#graph-depth").fill("3")
            page.locator("#graph-depth").dispatch_event("change")
            page.locator("#graph-options").evaluate("e=>e.open=false")
            page.wait_for_timeout(500)
            page.wait_for_function(
                "()=>!document.querySelector('#graph-depth-help').textContent.match(/Loading|Chargement/)"
            )
            page.locator("#graph-focus").evaluate("e=>e.click()")
            capture("neighborhood-depth3", width, height)
            if "fpna-20mb" in args.workbook.name:
                assert page.locator("#graph-canvas").evaluate(
                    "e=>e._flow.getViewport().zoom>=.35"
                ), "Large neighborhood cards are unreadable after fit"
            page.locator("#graph-depth").evaluate(
                "e=>{e.value=2;e.dispatchEvent(new Event('change'))}"
            )
            page.wait_for_timeout(500)
            capture("neighborhood-depth2", width, height)
            if "fpna-20mb" not in args.workbook.name:
                page.locator("#graph-search").fill("Summary!C10")
                page.locator('#graph-results [data-related="Summary!C10"]').click()
                for d in [1, 2, 3]:
                    page.locator("#graph-depth").evaluate(
                        "(e,d)=>{e.value=d;e.dispatchEvent(new Event('change'))}", d
                    )
                    page.wait_for_timeout(500)
                    capture(f"neighborhood-c10-depth{d}", width, height)
                page.locator("#graph-search").fill(target)
                page.locator(f'#graph-results [data-related="{target}"]').click()
            page.locator("#graph-depth").evaluate(
                "e=>{e.value=3;e.dispatchEvent(new Event('change'))}"
            )
            page.wait_for_timeout(500)
            aggregate = page.locator("#graph-canvas").evaluate(
                "e=>e._flow.getNodes().find(n=>n.data.original.aggregate)?.id"
            )
            if aggregate:
                page.locator("#graph-canvas").evaluate(
                    "(e,id)=>e._flow.select(id)", aggregate
                )
                page.wait_for_timeout(350)
                capture("neighborhood-group-cells", width, height)
                page.locator("#graph-neighborhood").click()
                page.wait_for_timeout(350)
            page.locator("#graph-fullscreen").click()
            capture("focus", width, height)
            page.keyboard.press("Escape")
            assert page.locator("#view-graph").get_attribute("aria-modal") is None
            page.locator("#graph-search").fill(target.split("!")[0])
            page.locator("#graph-results [data-related]").first.wait_for()
            capture("graph-search", width, height)
            page.keyboard.press("Escape")
            assert not page.locator("#graph-results").inner_text()
            page.locator("#graph-search").fill("")
            manifest["calculationTasks"] = api("projects/" + project)["tasks"]
            page.locator("#graph-build").evaluate("e=>e.click()")
            page.wait_for_function(
                "()=>document.querySelectorAll('#graph-breadcrumbs button').length>0"
            )
        browser.close()
    manifest["status"] = (
        "passed"
        if not manifest["errors"]
        and not any(m["overflow"] for m in manifest["metrics"])
        else "failed"
    )
    (output / "validation.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(output)
    print(json.dumps({"status": manifest["status"], "errors": manifest["errors"]}))
    return int(manifest["status"] != "passed")


if __name__ == "__main__":
    raise SystemExit(main())
