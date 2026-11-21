"""Saved capture loading stays local, recoverable and inert across navigation."""

import json
import os
from pathlib import Path
from urllib.parse import urlsplit

import pytest
from graph_test_helpers import close_workspace_menu, french_words, open_workspace_menu

playwright = pytest.importorskip("playwright.sync_api")
APP = Path(__file__).parents[1] / "src/linexcel/assets/app.html"


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


def _capture(task_id, **extra):
    return {
        "id": task_id,
        "operation": "capture",
        "status": "succeeded",
        "resultAvailable": True,
        **extra,
    }


def _images(sheet):
    return {"screenshots": [{"sheet": sheet, "url": f"images/{sheet}.svg"}]}


def _route_app(page, projects, tasks_by_project, task_result, config=None):
    requests, errors = [], []
    page.on("pageerror", lambda error: errors.append(str(error)))

    def handle(route):
        if (
            route.request.method == "POST"
            and route.request.url.endswith("/tasks")
            and route.request.post_data_json.get("operation") == "build_graph"
        ):
            route.fulfill(
                status=409,
                json={"error": {"message": "Overview unavailable in fixture"}},
            )
            return
        url = urlsplit(route.request.url)
        if url.path == "/mounted/":
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if url.path.endswith(("app-graph.js", "app-graph.css")):
            route.fulfill(
                body=(APP.parent / route.request.url.rsplit("/", 1)[-1]).read_text(
                    encoding="utf-8"
                ),
                content_type=(
                    "text/css"
                    if route.request.url.endswith(".css")
                    else "application/javascript"
                ),
            )
            return
        if "/images/" in url.path:
            if url.path.endswith("/Missing.svg"):
                route.fulfill(status=404, body="Image not found")
                return
            route.fulfill(
                body=(
                    "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100'/>"
                ),
                content_type="image/svg+xml",
            )
            return
        path = url.path.removeprefix("/mounted/api/")
        requests.append(path)
        assert route.request.method == "GET"  # Tab navigation never starts work.
        if path == "config":
            data = config or {}
        elif path == "projects":
            data = {"projects": projects}
        elif path.startswith("projects/"):
            project = next(p for p in projects if p["id"] == path.split("/")[1])
            data = {
                "project": project,
                "tasks": tasks_by_project[project["id"]],
                "graph": {
                    "nodes": [
                        {"id": "Data!A1", "sheet": "Data", "value": 42, "kind": "input"}
                    ],
                    "edges": [],
                    "meta": {"sheets": ["Data"]},
                },
            }
        elif path.startswith("tasks/"):
            task_result(route, path.split("/")[1])
            return
        else:
            raise AssertionError(path)
        route.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", handle)
    page.goto("http://localhost/mounted/")
    open_workspace_menu(page)
    page.locator(f"[data-project='{projects[0]['id']}']").click()
    page.locator("#project-title").get_by_text(
        projects[0]["name"], exact=True
    ).wait_for()
    return requests, errors


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_capture_tab_renders_while_loading_and_finds_older_nonempty_result(
    browser, viewport
):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    pending = []
    tasks = [_capture("old", createdAt=1), _capture("empty", createdAt=2)]

    def result(route, task_id):
        if task_id == "empty":
            pending.append(route)
        else:
            route.fulfill(json={"task": {**tasks[0], "result": _images("Older")}})

    try:
        requests, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": tasks}, result
        )
        page.locator("#tab-graph").click()
        page.locator("#tab-captures").click()
        page.locator("#capture-load-status").wait_for()
        assert page.locator("#view-captures").is_visible()
        assert page.locator("#view-graph").is_hidden()
        assert page.locator("#tab-captures").get_attribute("aria-selected") == "true"
        assert len(pending) == 1
        pending[0].fulfill(json={"task": {**tasks[1], "result": {"screenshots": []}}})
        page.locator("#capture-selected h3").get_by_text("Older", exact=True).wait_for()
        assert requests[-2:] == ["tasks/empty", "tasks/old"]
        assert (
            "dernières images disponibles"
            in page.locator("#capture-results").inner_text()
        )
        assert page.locator("#capture-load-status").count() == 0
        assert page.locator("#capture-selected img").is_visible()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("malformed", [None, "not an array", {"sheet": "Data"}])
def test_malformed_saved_capture_result_falls_back_without_breaking_tabs(
    browser, malformed
):
    page = browser.new_page()
    tasks = [_capture("old", createdAt=1), _capture("bad", createdAt=2)]

    def result(route, task_id):
        task = next(t for t in tasks if t["id"] == task_id)
        route.fulfill(
            json={
                "task": {
                    **task,
                    "result": {"screenshots": malformed}
                    if task_id == "bad"
                    else _images("Older"),
                }
            }
        )

    try:
        _, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": tasks}, result
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-selected h3").get_by_text("Older", exact=True).wait_for()
        assert "illisible" in page.locator("#capture-load-error").inner_text()
        assert page.locator("#capture-retry").is_visible()
        page.locator("#tab-graph").click()
        assert page.locator("#view-graph").is_visible()
        assert page.locator("#capture-load-error").is_hidden()
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
@pytest.mark.parametrize("navigate", [False, True])
def test_capture_result_failure_is_local_and_retryable_after_navigation(
    browser, viewport, navigate
):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    pending, attempts = [], []
    task = _capture("saved")

    def result(route, task_id):
        attempts.append(task_id)
        if len(attempts) == 1:
            pending.append(route)
        else:
            route.fulfill(json={"task": {**task, "result": _images("Recovered")}})

    try:
        _, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": [task]}, result
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-load-status").wait_for()
        if navigate:
            page.locator("#tab-graph").click()
        pending[0].fulfill(
            status=500, json={"error": {"message": "Capture unavailable"}}
        )
        page.wait_for_function("document.querySelector('#capture-load-error') !== null")
        assert page.locator("#message").inner_text() == ""
        if navigate:
            assert page.locator("#view-graph").is_visible()
            assert page.locator("#capture-load-error").is_hidden()
            page.locator("#tab-captures").click()
        else:
            assert page.locator("#view-captures").is_visible()
            assert page.locator("#capture-load-error").is_visible()
            page.locator("#capture-retry").click()
        page.locator("#capture-selected h3").get_by_text(
            "Recovered", exact=True
        ).wait_for()
        assert attempts == ["saved", "saved"]
        assert page.locator("#capture-load-error").count() == 0
        assert errors == []
    finally:
        page.close()


def test_pending_capture_from_another_project_cannot_replace_current_gallery(browser):
    page = browser.new_page()
    pending = []
    old = _capture("old")
    new = _capture("new", result=_images("Current"))

    def result(route, task_id):
        assert task_id == "old"
        pending.append(route)

    try:
        _, errors = _route_app(
            page,
            [{"id": "p", "name": "Old workbook"}, {"id": "p2", "name": "New workbook"}],
            {"p": [old], "p2": [new]},
            result,
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-load-status").wait_for()
        open_workspace_menu(page)
        page.locator("[data-project='p2']").click()
        page.locator("#capture-selected h3").get_by_text(
            "Current", exact=True
        ).wait_for()
        pending[0].fulfill(json={"task": {**old, "result": _images("Stale")}})
        page.wait_for_function(
            "document.querySelector('#capture-load-status') === null"
        )
        assert page.locator("#capture-selected h3").first.inner_text() == "Current"
        assert "Stale" not in page.locator("#capture-results").inner_text()
        assert errors == []
    finally:
        page.close()


def test_invalid_image_entries_and_base64_do_not_interrupt_application(browser):
    page = browser.new_page()
    task = _capture(
        "bad-image",
        result={
            "screenshots": [
                {"sheet": "Broken", "data": "data:image/png;base64,%%%%"},
                None,
            ]
        },
    )

    def result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": [task]}, result
        )
        page.locator("#tab-captures").click()
        assert "Image indisponible" in page.locator("#capture-selected").inner_text()
        page.locator("#capture-next").click()
        assert "Capture 2 / 2" in page.locator("#capture-selected").inner_text()
        assert "Image indisponible" in page.locator("#capture-selected").inner_text()
        page.locator("#tab-graph").click()
        assert page.locator("#view-graph").is_visible()
        assert errors == []
    finally:
        page.close()


def test_missing_capture_image_has_visible_error_and_keeps_navigation(browser):
    page = browser.new_page()
    task = _capture("missing-image", result=_images("Missing"))

    def result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": [task]}, result
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-image-error").wait_for()
        assert "Image indisponible" in page.locator("#capture-image-error").inner_text()
        assert page.locator("#capture-selected figure img").is_hidden()
        page.locator("#tab-graph").click()
        assert page.locator("#view-graph").is_visible()
        assert page.locator("#capture-image-error").is_hidden()
        assert errors == []
    finally:
        page.close()


def test_capture_search_reads_only_one_cache_sized_batch_then_can_continue(browser):
    page = browser.new_page()
    tasks = [_capture(str(i), createdAt=i) for i in range(13)]

    def result(route, task_id):
        task = next(t for t in tasks if t["id"] == task_id)
        route.fulfill(
            json={
                "task": {
                    **task,
                    "result": _images("Oldest")
                    if task_id == "0"
                    else {"screenshots": []},
                }
            }
        )

    try:
        requests, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": tasks}, result
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-older").wait_for()
        assert len([r for r in requests if r.startswith("tasks/")]) == 12
        page.locator("#capture-older").click()
        page.locator("#capture-selected h3").get_by_text(
            "Oldest", exact=True
        ).wait_for()
        assert len([r for r in requests if r.startswith("tasks/")]) == 13
        page.locator("#tab-graph").click()
        page.locator("#tab-captures").click()
        page.locator("#capture-selected h3").get_by_text(
            "Oldest", exact=True
        ).wait_for()
        assert len([r for r in requests if r.startswith("tasks/")]) == 13
        assert page.locator("#capture-older").count() == 0
        assert errors == []
    finally:
        page.close()


def test_new_completed_capture_is_loaded_after_cached_gallery_reentry(browser):
    page = browser.new_page()
    old = _capture("old", createdAt=1, result=_images("Old gallery"))
    queued = {"id": "new", "operation": "capture", "status": "queued", "createdAt": 2}
    saved = {**queued, "status": "succeeded", "resultAvailable": True, "finishedAt": 3}
    reads = []

    def result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page, [{"id": "p", "name": "Workbook"}], {"p": [old]}, result
        )
        page.locator("#tab-captures").click()
        page.locator("#capture-selected h3").get_by_text(
            "Old gallery", exact=True
        ).wait_for()
        page.route(
            "**/mounted/api/projects/p/tasks",
            lambda route: route.fulfill(json={"task": queued}),
        )

        def read_new(route):
            reads.append(route.request.method)
            task = (
                saved
                if len(reads) == 1
                else {**saved, "result": _images("New gallery")}
            )
            route.fulfill(json={"task": task})

        page.route("**/mounted/api/tasks/new", read_new)
        page.locator("#capture").click()
        playwright.expect(page.locator("#capture")).to_be_disabled()
        playwright.expect(page.locator("#capture")).to_be_enabled()
        assert reads == ["GET"]  # Completion metadata contains no image result.
        page.locator("#tab-graph").click()
        page.locator("#tab-captures").click()
        page.locator("#capture-selected h3").get_by_text(
            "New gallery", exact=True
        ).wait_for()
        assert reads == ["GET", "GET"]
        assert errors == []
    finally:
        page.close()


def _limits(server_memory, default_memory):
    return {
        "limits": {
            "memoryMb": server_memory,
            "uploadMb": 256,
            "storageMb": 150000,
            "storageUsedMb": 161,
            "maxSeconds": 3600,
            "defaults": {
                "longSeconds": 3600,
                "seconds": 120,
                "memoryMb": default_memory,
            },
        }
    }


@pytest.mark.parametrize(
    ("server_memory", "initial_memory", "user_memory"),
    [(256, 256, 256), (768, 768, 512), (16384, 8192, 512)],
)
def test_capture_uses_user_memory_without_raising_it_after_server_default(
    browser, server_memory, initial_memory, user_memory
):
    page = browser.new_page()
    submitted = []

    def unexpected_result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page,
            [{"id": "p", "name": "Memory budget"}],
            {"p": []},
            unexpected_result,
            config=_limits(server_memory, initial_memory),
        )
        assert page.locator("#memory").input_value() == str(initial_memory)
        assert page.locator("#memory").get_attribute("max") == str(server_memory)
        open_workspace_menu(page)
        page.get_by_text("Limites par opération", exact=True).click()
        page.locator("#memory").fill(str(user_memory))
        close_workspace_menu(page)

        def capture(route):
            submitted.append(route.request.post_data_json)
            route.fulfill(
                json={
                    "task": {
                        "id": "capture-user-budget",
                        "operation": "capture",
                        "status": "succeeded",
                        "result": {"screenshots": []},
                    }
                }
            )

        page.route("**/mounted/api/projects/p/tasks", capture)
        page.locator("#tab-captures").click()
        page.locator("#capture").click()
        playwright.expect(page.locator("#capture-results")).to_contain_text(
            "aucune image"
        )
        assert len(submitted) == 1
        assert submitted[0]["operation"] == "capture"
        assert submitted[0]["budget"]["memoryMb"] == user_memory
        assert page.locator("#memory").input_value() == str(user_memory)
        assert errors == []
    finally:
        page.close()


def test_limit_inputs_show_server_defaults_and_capture_runs_long(browser):
    page = browser.new_page()
    submitted = []

    def unexpected_result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page,
            [{"id": "p", "name": "Large workbook"}],
            {"p": []},
            unexpected_result,
            config=_limits(16384, 8192),
        )
        assert page.locator("#import-seconds").input_value() == "3600"
        assert page.locator("#seconds").input_value() == "120"
        assert page.locator("#memory").input_value() == "8192"
        assert page.locator("#memory").get_attribute("max") == "16384"
        for limit in ("#import-seconds", "#seconds"):
            assert page.locator(limit).get_attribute("max") == "3600"

        def capture(route):
            submitted.append(route.request.post_data_json)
            route.fulfill(
                json={
                    "task": {
                        "id": "capture-long",
                        "operation": "capture",
                        "status": "succeeded",
                        "result": {"screenshots": []},
                    }
                }
            )

        page.route("**/mounted/api/projects/p/tasks", capture)
        page.locator("#tab-captures").click()
        page.locator("#capture").click()
        playwright.expect(page.locator("#capture-results")).to_contain_text(
            "aucune image"
        )
        assert submitted[0]["budget"] == {"seconds": 3600, "memoryMb": 8192}
        page.locator("#tab-diagnostics").click()
        playwright.expect(page.locator("#runtime-limits")).to_contain_text(
            "161 / 150000 MiB"
        )
        assert errors == []
    finally:
        page.close()


def test_limit_failures_read_in_english(browser):
    from linexcel.web import WebConfig
    from linexcel.web_limits import (
        STORAGE_REFUSED,
        STORAGE_TRANSFER,
        limit_message,
        storage_quota_message,
    )

    config = WebConfig(Path("unused"), total_memory_mb=16384, max_storage_mb=4096)
    failures = [
        limit_message("timeout", config, operation="capture", seconds=3600),
        limit_message("timeout", config, operation="import", seconds=900),
        limit_message("timeout", config, operation="evaluate", seconds=120),
        limit_message("timeout", config, operation="document", seconds=600),
        limit_message("memory_limit", config, memory_mb=4608),
        limit_message("memory_limit", config, memory_mb=16384),
        limit_message("storage_limit", config),
        storage_quota_message(4096, STORAGE_REFUSED),
        storage_quota_message(4096, STORAGE_TRANSFER),
    ]
    tasks = [
        {
            "id": f"failed-{index}",
            "operation": "capture",
            "status": "failed",
            "error": {"kind": "limit", "message": message},
        }
        for index, message in enumerate(failures)
    ]
    page = browser.new_page()

    def unexpected_result(route, task_id):
        raise AssertionError(task_id)

    try:
        _, errors = _route_app(
            page,
            [{"id": "p", "name": "Limits"}],
            {"p": tasks},
            unexpected_result,
            config=_limits(16384, 8192),
        )
        page.locator("#ui-language").select_option("en")
        page.locator("#tab-tasks").click()
        notices = page.locator("#task-list .notice.error")
        playwright.expect(notices).to_have_count(len(failures))
        text = "\n".join(notices.all_inner_texts())
        for english in (
            "Time limit of 3600 s reached, the longest allowed.",
            "Raise “Long operations” in Settings › Operation limits (up to 60 min)",
            "Raise “Targeted recalculation” in Settings › Operation limits "
            "(up to 3600 s)",
            "Raise “Maximum duration” in Settings › AI documentation (up to 3600 s)",
            "Memory budget of 4.5 GB (4608 MiB) reached. Raise “Memory per operation”",
            "up to the server ceiling of 16 GB",
            "Memory budget of 16 GB (16384 MiB) reached, equal to the server ceiling.",
            "a --memory-mb value larger than 16384",
            "(4096 MiB, workbooks, indexes and views included): the operation was "
            "stopped and its result was not kept.",
            "the operation could not start.",
            "the transfer was interrupted.",
            "a larger --max-storage-mb value",
        ):
            assert english in text, english
        assert french_words(text) == [], text
        assert errors == []
    finally:
        page.close()
