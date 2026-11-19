#!/usr/bin/env python3
"""Launch, check, drive and tear down a disposable linexcel for verification.

Run from the repository root, always through uv so the working tree is the
code under test:

    uv run --extra screenshots python \
        .agents/skills/verify-linexcel/scripts/control_linexcel.py launch

Commands (each prints what it did and exits non-zero on failure):

    launch   build the synthetic workbooks, start `linexcel serve` on a free
             loopback port with a private data directory, wait until usable
    doctor   read-only: is that instance alive, ours, and running this tree?
    cli      run one `linexcel ...` command and record stdout/stderr/exit code
    graph    read a graph JSON written by `analyze --json`: totals, one node
    browser  one Chromium session against the app or a report file, as a
             sequence of --step actions; records screenshots and ARIA text

A browser step is a verb, then its arguments separated by " :: ":

    click SELECTOR                  fill SELECTOR :: TEXT
    press KEY | press SELECTOR :: KEY           select SELECTOR :: VALUE
    upload SELECTOR :: FILE         wait SELECTOR | wait-hidden SELECTOR
    wait-tasks  (app only: no server task queued or running)
    expect-text SELECTOR :: TEXT    expect-value SELECTOR :: VALUE
    expect-count SELECTOR :: N      expect-hidden SELECTOR
    shot NAME | shot-full NAME      aria NAME [:: SELECTOR]
    text NAME :: SELECTOR           download FILENAME :: SELECTOR

Selectors are Playwright selectors (CSS, or role=tab[name='Graph']). A failed
step stops the session, saves FAILED.png and exits 1. So do a JavaScript error
on the page and any request to a non-loopback address.
    cleanup  stop the server this run started and delete its scratch state;
             the evidence directory is kept

Everything lives under validation_screenshots/verify-linexcel/<run>/, which
git ignores: `scratch/` (data directory, fixtures, server log, browser state)
and `evidence/` (the proof). Nothing here reads or writes ~/.linexcel.

Placeholders usable in `cli` arguments and `browser` step arguments:
{fixtures}, {evidence}, {scratch}, {sales}, {stress}, {macros}.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shlex
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
RUNS = ROOT / "validation_screenshots" / "verify-linexcel"
VIEWPORTS = {"desktop": (1440, 900), "mobile": (390, 844)}
SALES = "fy26_commercial_review.xlsx"
STRESS = "validation_stress.xlsx"


class ControlError(Exception):
    """A verification step that did not hold; printed without a traceback."""


def run_dir(run: str) -> Path:
    if not run or any(c in run for c in "/\\:") or run.startswith("."):
        raise ControlError(f"invalid run id: {run!r}")
    return RUNS / run


def paths(run: str) -> dict[str, Path]:
    base = run_dir(run)
    scratch = base / "scratch"
    return {
        "base": base,
        "scratch": scratch,
        "evidence": base / "evidence",
        "fixtures": scratch / "fixtures",
        "data": scratch / "data",
        "instance": scratch / "instance.json",
        "state": scratch / "browser-state.json",
    }


def placeholders(run: str) -> dict[str, str]:
    p = paths(run)
    return {
        "fixtures": str(p["fixtures"]),
        "evidence": str(p["evidence"]),
        "scratch": str(p["scratch"]),
        "sales": str(p["fixtures"] / SALES),
        "stress": str(p["fixtures"] / STRESS),
        "macros": str(ROOT / "tests" / "fixtures" / "macros.xlsm"),
    }


def expand(run: str, text: str) -> str:
    for key, value in placeholders(run).items():
        text = text.replace("{" + key + "}", value)
    return text


def source_fingerprint() -> str:
    """Hash of every file the running server loaded from this checkout."""
    digest = hashlib.sha256()
    for path in sorted((ROOT / "src" / "linexcel").rglob("*")):
        if path.is_file() and "__pycache__" not in path.parts:
            digest.update(path.relative_to(ROOT).as_posix().encode())
            digest.update(path.read_bytes().replace(b"\r\n", b"\n"))
    return digest.hexdigest()


def git_head() -> str:
    out = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=ROOT, capture_output=True, text=True
    )
    return out.stdout.strip() or "unknown"


def process_alive(pid: int) -> bool:
    if sys.platform == "win32":
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
        )
        return f'"{pid}"' in out.stdout
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def stop_tree(pid: int) -> None:
    """Stop one pid and its descendants. Callers pass only pids they started."""
    if not process_alive(pid):
        return
    if sys.platform == "win32":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True)
    else:
        import signal

        try:
            os.killpg(os.getpgid(pid), signal.SIGINT)
        except OSError:
            return
    deadline = time.monotonic() + 15
    while process_alive(pid) and time.monotonic() < deadline:
        time.sleep(0.2)
    if process_alive(pid) and sys.platform != "win32":
        import signal

        os.killpg(os.getpgid(pid), signal.SIGKILL)
        time.sleep(1)


def http_json(url: str, timeout: float = 5) -> tuple[int, dict | None]:
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, json.loads(response.read())
    except urllib.error.HTTPError as exc:
        return exc.code, None
    except (urllib.error.URLError, OSError, ValueError):
        return 0, None


def read_instance(run: str) -> dict:
    path = paths(run)["instance"]
    if not path.exists():
        raise ControlError(
            f"no instance recorded for run {run!r}; run `launch --run {run}` first"
        )
    return json.loads(path.read_text("utf-8"))


def linexcel(*args: str, **kwargs) -> subprocess.CompletedProcess:
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, "-m", "linexcel", *args],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        **kwargs,
    )


# --------------------------------------------------------------------- launch


def cmd_launch(args) -> int:
    run = args.run or time.strftime("%Y%m%d-%H%M%S")
    p = paths(run)
    if p["scratch"].exists():
        raise ControlError(
            f"run {run!r} already has scratch state; `cleanup --run {run}` first"
        )
    for key in ("fixtures", "data", "evidence"):
        p[key].mkdir(parents=True, exist_ok=True)
    # `linexcel analyze -o/--json` does not create parent directories.
    (p["evidence"] / "reports").mkdir(exist_ok=True)

    sys.path.insert(0, str(ROOT))
    import validation_workbooks

    (p["fixtures"] / SALES).write_bytes(validation_workbooks.build_sales_workbook())
    (p["fixtures"] / STRESS).write_bytes(validation_workbooks.build_stress_workbook())

    log = (p["scratch"] / "serve.log").open("w", encoding="utf-8")
    flags = 0
    if sys.platform == "win32":
        flags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
    command = [
        sys.executable,
        "-m",
        "linexcel",
        "serve",
        "--port",
        "0",
        "--data-dir",
        str(p["data"]),
        *args.serve_args,
    ]
    server = subprocess.Popen(
        command,
        cwd=ROOT,
        stdout=log,
        stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        creationflags=flags,
        start_new_session=sys.platform != "win32",
        env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1"),
    )
    # The data directory is new and private to this run, so whatever server
    # publishes itself there is the one started above. Its pid can differ from
    # the launcher's: a Windows venv `python.exe` is a trampoline that starts
    # the real interpreter as a child.
    published = p["data"] / ".access" / "server.json"
    deadline = time.monotonic() + args.timeout
    record = None
    while time.monotonic() < deadline:
        if server.poll() is not None:
            raise ControlError(
                f"linexcel serve exited with {server.returncode}; see "
                f"{p['scratch'] / 'serve.log'}"
            )
        if published.exists():
            try:
                candidate = json.loads(published.read_text("utf-8"))
            except ValueError:
                candidate = None
            if candidate and process_alive(candidate.get("pid", 0)):
                status, config = http_json(candidate["url"] + "api/config")
                # Usable means the app answers with its configuration, not
                # merely that the port accepts a connection.
                if status == 200 and config and "limits" in config:
                    record = candidate
                    break
        time.sleep(0.2)
    if record is None:
        stop_tree(server.pid)
        raise ControlError(f"linexcel serve was not usable after {args.timeout}s")

    version = linexcel("--version").stdout.strip()
    instance = {
        "run": run,
        "pid": record["pid"],
        "launcherPid": server.pid,
        "url": record["url"],
        "publication": record["publication"],
        "command": command,
        "version": version,
        "head": git_head(),
        "sourceFingerprint": source_fingerprint(),
        "startedAt": record["startedAt"],
    }
    p["instance"].write_text(json.dumps(instance, indent=2), encoding="utf-8")
    (p["evidence"] / "run.json").write_text(
        json.dumps(
            {k: instance[k] for k in ("run", "url", "command", "version", "head")},
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"run:      {run}")
    print(f"ready:    {record['url']}  (pid {record['pid']}, {version})")
    print(f"fixtures: {p['fixtures']}")
    print(f"evidence: {p['evidence']}")
    return 0


# --------------------------------------------------------------------- doctor


def doctor_report(run: str) -> list[tuple[str, bool, str]]:
    p = paths(run)
    instance = read_instance(run)
    checks: list[tuple[str, bool, str]] = []
    alive = process_alive(instance["pid"])
    checks.append(("process", alive, f"pid {instance['pid']}"))
    published = p["data"] / ".access" / "server.json"
    try:
        record = json.loads(published.read_text("utf-8"))
    except (OSError, ValueError):
        record = {}
    ours = (
        record.get("pid") == instance["pid"]
        and record.get("publication") == instance["publication"]
        and record.get("url") == instance["url"]
    )
    checks.append(
        ("ownership", ours, "data directory still published by the pid we started")
    )
    status, config = http_json(instance["url"] + "api/config")
    local = bool(config and config.get("session", {}).get("localProfile"))
    checks.append(
        ("http", status == 200 and local, f"{instance['url']}api/config -> {status}")
    )
    # Minting a session link would write a token; doctor only reads. A profile
    # beside the publication is what `linexcel session-link` needs to succeed.
    profile = p["data"] / ".access" / "profile.json"
    checks.append(
        ("session", profile.is_file(), "local profile present for session links")
    )
    fresh = source_fingerprint() == instance["sourceFingerprint"]
    checks.append(
        (
            "build",
            fresh,
            "src/linexcel unchanged since launch"
            if fresh
            else "src/linexcel changed since launch: the server runs old code; "
            "cleanup and launch again",
        )
    )
    for name in (SALES, STRESS):
        checks.append((f"fixture:{name}", (p["fixtures"] / name).is_file(), ""))
    return checks


def cmd_doctor(args) -> int:
    checks = doctor_report(args.run)
    for name, ok, detail in checks:
        print(f"{'ok  ' if ok else 'FAIL'} {name:<32} {detail}")
    healthy = all(ok for _, ok, _ in checks)
    print("healthy: drive it" if healthy else "unhealthy: do not drive this run")
    return 0 if healthy else 1


# ------------------------------------------------------------------------ cli


def cmd_cli(args) -> int:
    p = paths(args.run)
    if not p["base"].exists():
        raise ControlError(f"unknown run {args.run!r}; run `launch` first")
    arguments = [expand(args.run, a) for a in args.arguments]
    if arguments[:1] == ["--"]:
        arguments = arguments[1:]
    if not arguments:
        raise ControlError("give the linexcel arguments after `--`")
    if arguments[0] in {"serve"}:
        raise ControlError("start servers with `launch`, which records what to stop")
    if arguments[0] in {"library", "doctor", "session-link"} and not any(
        a == "--data-dir" or a.startswith("--data-dir=") for a in arguments
    ):
        # The default data directory is the user's real ~/.linexcel.
        arguments[1:1] = ["--data-dir", str(p["data"])]
    out = p["evidence"] / "cli"
    out.mkdir(parents=True, exist_ok=True)
    label = args.name or f"{len(list(out.glob('*.json'))) + 1:02d}-{arguments[0]}"
    started = time.time()
    done = linexcel(*arguments, timeout=args.timeout)
    record = {
        "command": "linexcel " + " ".join(shlex.quote(a) for a in arguments),
        "exitCode": done.returncode,
        "seconds": round(time.time() - started, 2),
        "stdoutFile": f"{label}.stdout.txt",
        "stderrFile": f"{label}.stderr.txt",
    }
    (out / f"{label}.stdout.txt").write_text(done.stdout, encoding="utf-8")
    (out / f"{label}.stderr.txt").write_text(done.stderr, encoding="utf-8")
    (out / f"{label}.json").write_text(json.dumps(record, indent=2), encoding="utf-8")
    print(f"$ {record['command']}")
    print(f"exit {done.returncode} in {record['seconds']}s")
    for name, text in (("stdout", done.stdout), ("stderr", done.stderr)):
        lines = text.splitlines()
        shown = lines if len(lines) <= args.tail else lines[-args.tail :]
        if len(lines) > len(shown):
            print(f"--- {name} (last {args.tail} of {len(lines)} lines)")
        else:
            print(f"--- {name}")
        for line in shown:
            print(line.encode("ascii", "replace").decode())
    print(f"recorded: {out / (label + '.json')}")
    if args.expect_exit is not None and done.returncode != args.expect_exit:
        raise ControlError(f"expected exit {args.expect_exit}, got {done.returncode}")
    return 0


# ---------------------------------------------------------------------- graph


def cmd_graph(args) -> int:
    """Read a graph JSON written by `linexcel analyze --json`: a second view."""
    p = paths(args.run)
    source = Path(expand(args.run, args.json))
    graph = json.loads(source.read_text("utf-8"))
    meta = graph.get("meta", {})
    report: dict = {
        "file": str(source),
        "executionStatus": meta.get("execution", {}).get("status"),
        "stats": {
            key: value
            for key, value in meta.get("stats", {}).items()
            if not isinstance(value, (list, dict))
        },
        "counts": {"nodes": len(graph["nodes"]), "edges": len(graph["edges"])},
    }
    if args.node:
        node = next((n for n in graph["nodes"] if n["id"] == args.node), None)
        if node is None:
            raise ControlError(f"node {args.node!r} is not in {source.name}")
        report["node"] = {
            key: node.get(key)
            for key in ("id", "kind", "sheet", "addr", "formula", "value")
        }
        report["node"]["valueSource"] = node.get("valueSource")
        report["precedents"] = sorted(
            e["source"] for e in graph["edges"] if e["target"] == args.node
        )
        report["dependents"] = sorted(
            e["target"] for e in graph["edges"] if e["source"] == args.node
        )
    text = json.dumps(report, ensure_ascii=False, indent=2)
    out = p["evidence"] / "cli"
    out.mkdir(parents=True, exist_ok=True)
    name = args.name or f"graph-{source.stem}"
    (out / f"{name}.json").write_text(text, encoding="utf-8")
    print(text.encode("ascii", "backslashreplace").decode())
    print(f"recorded: {out / (name + '.json')}")
    return 0


# -------------------------------------------------------------------- browser


def run_steps(page, steps, out: Path, run: str, timeout_ms: int, base: str | None):
    from playwright.sync_api import expect

    log = []
    for raw in steps:
        # "verb first argument :: second argument". No nested quoting, so a
        # step reads the same in PowerShell and in a POSIX shell.
        verb, _, tail = raw.strip().partition(" ")
        rest = [expand(run, a.strip()) for a in tail.split(" :: ")] if tail else []
        entry = {"step": raw, "status": "ok"}
        started = time.time()
        try:
            if verb == "click":
                page.locator(rest[0]).first.click(timeout=timeout_ms)
            elif verb == "fill":
                page.locator(rest[0]).first.fill(rest[1], timeout=timeout_ms)
            elif verb == "press":
                if len(rest) == 1:
                    page.keyboard.press(rest[0])
                else:
                    page.locator(rest[0]).first.press(rest[1], timeout=timeout_ms)
            elif verb == "select":
                page.locator(rest[0]).first.select_option(rest[1], timeout=timeout_ms)
            elif verb == "upload":
                page.locator(rest[0]).first.set_input_files(rest[1:])
            elif verb == "wait":
                page.locator(rest[0]).first.wait_for(
                    state="visible", timeout=timeout_ms
                )
            elif verb == "wait-hidden":
                page.locator(rest[0]).first.wait_for(state="hidden", timeout=timeout_ms)
            elif verb == "wait-ms":
                page.wait_for_timeout(int(rest[0]))
            elif verb == "wait-tasks":
                if base is None:
                    raise ControlError("wait-tasks needs --open app")
                wait_tasks(page, base, timeout_ms)
            elif verb == "expect-text":
                expect(page.locator(rest[0]).first).to_contain_text(
                    rest[1], timeout=timeout_ms
                )
            elif verb == "expect-value":
                expect(page.locator(rest[0]).first).to_have_value(
                    rest[1], timeout=timeout_ms
                )
            elif verb == "expect-count":
                expect(page.locator(rest[0])).to_have_count(
                    int(rest[1]), timeout=timeout_ms
                )
            elif verb == "expect-hidden":
                expect(page.locator(rest[0]).first).to_be_hidden(timeout=timeout_ms)
            elif verb == "shot":
                page.screenshot(path=str(out / f"{rest[0]}.png"))
                entry["artifact"] = f"{rest[0]}.png"
            elif verb == "shot-full":
                page.screenshot(path=str(out / f"{rest[0]}.png"), full_page=True)
                entry["artifact"] = f"{rest[0]}.png"
            elif verb == "aria":
                selector = rest[1] if len(rest) > 1 else "body"
                text = page.locator(selector).first.aria_snapshot()
                (out / f"{rest[0]}.aria.txt").write_text(text, encoding="utf-8")
                entry["artifact"] = f"{rest[0]}.aria.txt"
            elif verb == "text":
                text = page.locator(rest[1]).first.inner_text(timeout=timeout_ms)
                (out / f"{rest[0]}.txt").write_text(text, encoding="utf-8")
                entry["artifact"] = f"{rest[0]}.txt"
            elif verb == "download":
                with page.expect_download(timeout=timeout_ms) as waiting:
                    page.locator(rest[1]).first.click(timeout=timeout_ms)
                waiting.value.save_as(str(out / rest[0]))
                entry["artifact"] = rest[0]
            else:
                raise ControlError(f"unknown step verb: {verb}")
        except Exception as exc:  # noqa: BLE001 - every step failure is evidence
            entry["status"] = "failed"
            entry["error"] = str(exc).splitlines()[0][:400]
            try:
                page.screenshot(path=str(out / "FAILED.png"))
                entry["artifact"] = "FAILED.png"
            except Exception:  # noqa: BLE001
                pass
        entry["seconds"] = round(time.time() - started, 2)
        log.append(entry)
        mark = "ok  " if entry["status"] == "ok" else "FAIL"
        note = entry.get("error") or entry.get("artifact") or ""
        print(f"{mark} {raw}  {note}".rstrip().encode("ascii", "replace").decode())
        if entry["status"] != "ok":
            break
    return log


def wait_tasks(page, base: str, timeout_ms: int) -> None:
    """Wait until the open project has no queued or running server task."""
    project = page.evaluate("localStorage.getItem('linexcel-project')")
    if not project:
        raise ControlError("no project is open in this browser")
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        response = page.request.get(f"{base}api/projects/{project}")
        if not response.ok:
            raise ControlError(f"project request failed: {response.status}")
        tasks = response.json().get("tasks", [])
        if not any(t["status"] in ("queued", "running") for t in tasks):
            return
        page.wait_for_timeout(300)
    raise ControlError("server tasks still queued or running")


def cmd_browser(args) -> int:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ControlError(
            "Playwright is missing: use `uv run --extra screenshots python ...` "
            "and `uv run --extra screenshots playwright install chromium`"
        ) from None

    p = paths(args.run)
    width, height = VIEWPORTS[args.viewport]
    out = p["evidence"] / "browser" / args.name
    out.mkdir(parents=True, exist_ok=True)
    base = None
    state = None
    if args.open == "app":
        unhealthy = [n for n, ok, _ in doctor_report(args.run) if not ok]
        if unhealthy:
            raise ControlError(f"doctor failed ({', '.join(unhealthy)}); not driving")
        base = read_instance(args.run)["url"]
        if p["state"].exists() and not args.fresh:
            state = str(p["state"])
            target = base + "?lang=en"
        else:
            link = linexcel("session-link", "--data-dir", str(p["data"]))
            if link.returncode != 0:
                raise ControlError(f"session-link failed: {link.stderr.strip()}")
            url, _, fragment = link.stdout.strip().partition("#")
            target = f"{url}?lang=en#{fragment}"
    else:
        report = Path(expand(args.run, args.open))
        if not report.is_file():
            raise ControlError(f"report not found: {report}")
        target = report.resolve().as_uri()

    page_errors: list[str] = []
    external: list[str] = []
    with sync_playwright() as driver:
        browser = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        context = browser.new_context(
            viewport={"width": width, "height": height},
            storage_state=state,
            accept_downloads=True,
        )
        page = context.new_page()
        page.on("pageerror", lambda error: page_errors.append(str(error)))

        def watch(request):
            url = request.url
            if not url.startswith(
                ("http://127.0.0.1", "http://localhost", "file:", "data:", "blob:")
            ):
                external.append(url)

        page.on("request", watch)
        page.goto(target)
        if base is not None:
            # Attached, not visible: on a narrow screen with workspaces already
            # imported, a new browser lands on the workspace picker and the
            # connection status is not displayed.
            page.locator("#connection").wait_for(
                state="attached", timeout=args.timeout * 1000
            )
            page.wait_for_function(
                "() => !document.querySelector('#session-required') || "
                "document.querySelector('#session-required').hidden",
                timeout=args.timeout * 1000,
            )
        log = run_steps(page, args.step, out, args.run, args.timeout * 1000, base)
        if base is not None:
            context.storage_state(path=str(p["state"]))
        browser.close()

    failed = [e for e in log if e["status"] != "ok"]
    record = {
        "open": args.open,
        "target": target.split("#")[0],
        "viewport": {"name": args.viewport, "width": width, "height": height},
        "steps": log,
        "pageErrors": page_errors,
        "externalRequests": external,
    }
    (out / "session.json").write_text(json.dumps(record, indent=2), encoding="utf-8")
    print(f"page errors: {len(page_errors)}  external requests: {len(external)}")
    for line in page_errors + external:
        print(f"  {line[:200]}".encode("ascii", "replace").decode())
    print(f"recorded: {out}")
    if failed:
        raise ControlError(f"step failed: {failed[0]['step']}")
    if page_errors and not args.allow_page_errors:
        raise ControlError("the page raised JavaScript errors")
    if external and not args.allow_external:
        raise ControlError("the page requested a non-loopback address")
    return 0


# -------------------------------------------------------------------- cleanup


def cmd_cleanup(args) -> int:
    p = paths(args.run)
    if not p["base"].exists():
        raise ControlError(f"unknown run {args.run!r}")
    if p["instance"].exists():
        instance = json.loads(p["instance"].read_text("utf-8"))
        pid = instance["pid"]
        published = p["data"] / ".access" / "server.json"
        try:
            record = json.loads(published.read_text("utf-8"))
        except (OSError, ValueError):
            record = {}
        ours = (
            record.get("pid") == pid
            and record.get("publication") == instance["publication"]
        )
        if process_alive(pid) and ours:
            # Stop exactly the pid this run started, and the analysis workers
            # it spawned; never by process name.
            stop_tree(pid)
            if process_alive(pid):
                raise ControlError(f"pid {pid} did not stop; scratch state kept")
            print(f"stopped:  pid {pid}")
        elif process_alive(pid):
            raise ControlError(
                f"pid {pid} is alive but no longer publishes this data directory; "
                "refusing to kill a process this run cannot prove it started"
            )
        else:
            print(f"stopped:  pid {pid} was already gone")
        stop_tree(instance.get("launcherPid", pid))
    if p["scratch"].exists():
        # The server stores tasks under owner/project/task ids, past the
        # 260-character Windows path limit; it writes them with an extended
        # prefix, and removing them needs the same prefix.
        target = str(p["scratch"].resolve())
        if sys.platform == "win32" and not target.startswith("\\\\?\\"):
            target = "\\\\?\\" + target
        error = None
        # Windows releases the handles of a stopped process a moment later.
        for _ in range(20):
            try:
                shutil.rmtree(target)
                error = None
                break
            except OSError as exc:
                error = exc
                time.sleep(0.5)
        if error is not None:
            raise ControlError(f"could not remove {p['scratch']}: {error}")
        print(f"removed:  {p['scratch']}")
    kept = sorted(x for x in p["evidence"].rglob("*") if x.is_file())
    print(f"evidence: {p['evidence']}  ({len(kept)} files kept)")
    return 0


# ----------------------------------------------------------------------- main


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    sub = parser.add_subparsers(dest="command", required=True)

    launch = sub.add_parser("launch", help="start a disposable instance")
    launch.add_argument("--run", help="run id (default: a timestamp)")
    launch.add_argument("--timeout", type=int, default=60)
    launch.add_argument(
        "serve_args",
        nargs=argparse.REMAINDER,
        help="extra `linexcel serve` options after `--`, such as --ai",
    )
    launch.set_defaults(func=cmd_launch)

    doctor = sub.add_parser("doctor", help="read-only health check of a run")
    doctor.add_argument("--run", required=True)
    doctor.set_defaults(func=cmd_doctor)

    cli = sub.add_parser("cli", help="run and record one linexcel command")
    cli.add_argument("--run", required=True)
    cli.add_argument("--name", help="evidence label (default: numbered)")
    cli.add_argument("--timeout", type=int, default=600)
    cli.add_argument("--tail", type=int, default=30, help="output lines to echo")
    cli.add_argument("--expect-exit", type=int, help="fail on another exit code")
    cli.add_argument("arguments", nargs=argparse.REMAINDER)
    cli.set_defaults(func=cmd_cli)

    graph = sub.add_parser("graph", help="summarise a graph JSON, or one node")
    graph.add_argument("--run", required=True)
    graph.add_argument("--json", required=True, help="file from analyze --json")
    graph.add_argument("--node", help="node id, such as c:Summary!C10")
    graph.add_argument("--name", help="evidence label")
    graph.set_defaults(func=cmd_graph)

    browser = sub.add_parser("browser", help="one recorded browser session")
    browser.add_argument("--run", required=True)
    browser.add_argument(
        "--open", required=True, help="`app`, or the path of a report HTML file"
    )
    browser.add_argument("--name", required=True, help="evidence folder name")
    browser.add_argument("--viewport", choices=sorted(VIEWPORTS), default="desktop")
    browser.add_argument("--timeout", type=int, default=30, help="seconds per step")
    browser.add_argument(
        "--fresh",
        action="store_true",
        help="ignore the saved browser state and enter through a new session link",
    )
    browser.add_argument("--allow-page-errors", action="store_true")
    browser.add_argument("--allow-external", action="store_true")
    browser.add_argument("--step", action="append", default=[], metavar="'VERB ARGS'")
    browser.set_defaults(func=cmd_browser)

    cleanup = sub.add_parser("cleanup", help="stop the instance, keep the evidence")
    cleanup.add_argument("--run", required=True)
    cleanup.set_defaults(func=cmd_cleanup)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if getattr(args, "serve_args", None) and args.serve_args[0] == "--":
        args.serve_args = args.serve_args[1:]
    try:
        return args.func(args)
    except ControlError as exc:
        print(f"control_linexcel: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
