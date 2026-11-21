"""Sidebar storage and limits line, and the settings dialog, as a user meets them."""

import json
import os
from pathlib import Path

import pytest
from graph_test_helpers import french_words, open_workspace_menu

playwright = pytest.importorskip("playwright.sync_api")
APP = Path(__file__).parents[1] / "src/linexcel/assets/app.html"
DESKTOP, PHONE = (1440, 900), (390, 844)


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


def _config(**limits):
    """A server with unusual numbers: 9 GiB ceiling, 25 min, 45 s, 4.5 GiB."""
    return {
        "session": {"localProfile": True, "required": False},
        "uploads": {"chunkBytes": 8388608, "localImport": True},
        "ai": {"available": False, "languages": ["en", "fr", "de"]},
        "limits": {
            "memoryMb": 9216,
            "uploadMb": 256,
            "retentionDays": None,
            "storageMb": 141312,
            "storageUsedMb": 1946,
            "maxSeconds": 3600,
            "defaults": {"longSeconds": 1500, "seconds": 45, "memoryMb": 4608},
            **limits,
        },
    }


class Server:
    """The application's HTTP API as a route: records what the page submits."""

    def __init__(self, page, config):
        self.page, self.config = page, config
        self.projects = [{"id": "p", "name": "Budget review", "revision": 1}]
        self.tasks, self.submitted, self.errors = {}, [], []
        self.task_status = "succeeded"
        page.on("pageerror", lambda error: self.errors.append(str(error)))
        page.route("**/*", self.handle)

    def json(self, route, body, status=200):
        route.fulfill(
            status=status, body=json.dumps(body), content_type="application/json"
        )

    def task(self, operation, status=None):
        task = {
            "id": f"t{len(self.tasks)}",
            "operation": operation,
            "status": status or self.task_status,
            "result": {"screenshots": []}
            if operation == "capture"
            else {"status": "completed", "value": 3},
        }
        self.tasks[task["id"]] = task
        return task

    def handle(self, route):
        request = route.request
        url = request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            name = url.rsplit("/", 1)[-1]
            route.fulfill(
                body=(APP.parent / name).read_text(encoding="utf-8"),
                content_type="text/css" if name.endswith(".css") else "text/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1].split("?", 1)[0]
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        path = path.removeprefix("api/")
        method = request.method
        if path == "config":
            return self.json(route, self.config)
        if path == "projects":
            return self.json(route, {"projects": self.projects})
        if path == "health":
            return self.json(
                route,
                {"checks": {"formualizer": {"status": "ready", "version": "0.14"}}},
            )
        if path == "projects/local" and method == "POST":
            body = request.post_data_json
            self.submitted.append({"operation": "import", **body})
            project = {"id": "p2", "name": body.get("name") or "Local", "revision": 1}
            self.projects.insert(0, project)
            return self.json(
                route,
                {"project": project, "task": self.task("import", status="running")},
            )
        if (
            path.startswith("projects/")
            and path.endswith("/tasks")
            and method == "POST"
        ):
            body = request.post_data_json
            if body["operation"] == "build_graph":
                return self.json(route, {"error": {"message": "No overview"}}, 409)
            self.submitted.append(body)
            return self.json(route, {"task": self.task(body["operation"])})
        if path.startswith("projects/") and path.count("/") == 1:
            project_id = path.split("/")[1]
            if method == "DELETE":
                self.projects = [p for p in self.projects if p["id"] != project_id]
                self.config["limits"]["storageUsedMb"] = 614
                return self.json(route, {"deleted": True})
            project = next(p for p in self.projects if p["id"] == project_id)
            tasks = [t for t in self.tasks.values() if t["operation"] == "import"]
            return self.json(
                route,
                {
                    "project": project,
                    "tasks": tasks,
                    "graph": {
                        "nodes": [
                            {
                                "id": "Data!A1",
                                "sheet": "Data",
                                "value": 1,
                                "kind": "input",
                            },
                            {
                                "id": "Data!B1",
                                "sheet": "Data",
                                "cell": "B1",
                                "formula": "=A1+2",
                                "kind": "formula",
                            },
                        ],
                        "edges": [{"source": "Data!A1", "target": "Data!B1"}],
                        "meta": {"sheets": ["Data"]},
                    },
                },
            )
        if path.startswith("tasks/"):
            return self.json(route, {"task": self.tasks[path.split("/")[1]]})
        return self.json(route, {})


def _open(browser, viewport=DESKTOP, lang="fr", config=None, project=True):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    server = Server(page, config or _config())
    page.goto(f"http://localhost/mounted/?lang={lang}")
    page.locator("#connection[data-state='ok']").wait_for(state="attached")
    if project:
        open_workspace_menu(page)
        page.locator("[data-project='p']").click()
        page.locator("#project-title").get_by_text("Budget review").wait_for()
    return page, server


def _open_limits_line(page):
    """Show the import form, where the limits line sits above "Importer"."""
    open_workspace_menu(page)
    if not page.locator("#import-panel").evaluate("e => e.open"):
        page.locator("#import-panel > summary").click()
    page.locator("#limits-edit").scroll_into_view_if_needed()


def _open_settings(page):
    open_workspace_menu(page)
    page.locator("#settings-open").click()
    dialog = page.get_by_role("dialog", name="Réglages", exact=True)
    playwright.expect(dialog).to_be_visible()
    return dialog


def _close_settings(page):
    page.locator("#settings-done").click()
    playwright.expect(page.locator("#settings")).to_be_hidden()


def _wait_for(page, condition, seconds=10):
    """Wait for what the server received; the page runs while Playwright waits."""
    for _ in range(seconds * 10):
        if condition():
            return
        page.wait_for_timeout(100)
    raise AssertionError("condition not met")


def _start_capture(page):
    page.locator("#tab-captures").click()
    page.locator("#capture").click()


def _select_formula(page):
    page.locator("#tab-nodes").click()
    page.locator("[data-node='Data!B1']").click()


def test_limit_fields_show_server_defaults_in_their_units_and_budgets_convert_them(
    browser,
):
    page, server = _open(browser)
    try:
        _open_limits_line(page)
        line = page.locator(".limits-line")
        playwright.expect(line).to_contain_text("Limites : 25 min · 45 s · 4,5 Go")
        page.locator("#limits-edit").click()
        assert page.locator("#limit-long-minutes").input_value() == "25"
        assert page.locator("#limit-targeted-seconds").input_value() == "45"
        assert page.locator("#limit-memory-gb").input_value() == "4.5"
        playwright.expect(page.locator("#settings-limits")).to_contain_text(
            "Plafond du serveur : 9 Go (option --memory-mb 9216). Conseillé : 4,5 Go."
        )
        _close_settings(page)
        _start_capture(page)
        _select_formula(page)  # The selection starts its calculation after 180 ms.
        _wait_for(page, lambda: len(server.submitted) == 2)
        budgets = {s["operation"]: s["budget"] for s in server.submitted}
        assert budgets == {
            "capture": {"seconds": 1500, "memoryMb": 4608},
            "evaluate": {"seconds": 45, "memoryMb": 4608},
        }

        _open_limits_line(page)
        page.locator("#limits-edit").click()
        page.locator("#limit-long-minutes").fill("40")
        page.locator("#limit-targeted-seconds").fill("90")
        page.locator("#limit-memory-gb").fill("6")
        _close_settings(page)
        playwright.expect(line).to_contain_text("Limites : 40 min · 90 s · 6 Go")
        page.locator("#import-source").select_option("local")
        page.locator("#local-path").fill("C:\\data\\big.xlsx")
        page.locator("#project-name").fill("Big")
        page.locator("#import-button").click()
        page.locator("#project-title").get_by_text("Big").wait_for()
        assert server.submitted[-1]["operation"] == "import"
        assert server.submitted[-1]["budget"] == {"seconds": 2400, "memoryMb": 6144}
        assert server.errors == []
    finally:
        page.close()


def test_advised_values_come_back_and_an_out_of_range_value_is_refused(browser):
    page, server = _open(browser)
    try:
        dialog = _open_settings(page)
        page.locator("#limit-long-minutes").fill("7")
        page.locator("#limit-memory-gb").fill("12")
        page.locator("#limit-memory-gb").press("Tab")
        error = dialog.locator("#limit-memory-gb-error")
        playwright.expect(error).to_have_text(
            "« Mémoire par opération » : choisissez une valeur de 0,5 à 9 Go "
            "dans Réglages › Limites par opération."
        )
        assert page.locator("#limit-memory-gb").get_attribute("aria-invalid") == "true"
        _close_settings(page)
        _start_capture(page)
        playwright.expect(page.locator("#message")).to_contain_text(
            "« Mémoire par opération » : choisissez une valeur de 0,5 à 9 Go"
        )
        assert server.submitted == []  # Refused in the page, never sent.

        _open_settings(page)
        page.get_by_role("button", name="Rétablir les valeurs conseillées").click()
        assert page.locator("#limit-long-minutes").input_value() == "25"
        assert page.locator("#limit-targeted-seconds").input_value() == "45"
        assert page.locator("#limit-memory-gb").input_value() == "4.5"
        playwright.expect(error).to_be_hidden()
        _close_settings(page)
        _start_capture(page)
        _wait_for(page, lambda: len(server.submitted) == 1)
        assert server.submitted[-1]["budget"] == {"seconds": 1500, "memoryMb": 4608}
        assert server.errors == []
    finally:
        page.close()


def test_storage_line_follows_an_import_that_ends_and_a_deletion(browser):
    page, server = _open(browser)
    try:
        storage = page.locator("#storage")
        playwright.expect(storage).to_contain_text("Stockage local")
        playwright.expect(storage).to_contain_text("1,9 Go sur 138 Go")
        meter = page.get_by_role("meter", name="Stockage local")
        assert meter.get_attribute("aria-valuetext") == "1,9 Go sur 138 Go"
        assert meter.get_attribute("aria-valuenow") == "1"

        # An import that ends grows the storage; the page follows on its own.
        _open_limits_line(page)
        page.locator("#import-source").select_option("local")
        page.locator("#local-path").fill("C:\\data\\big.xlsx")
        page.locator("#project-name").fill("Big")
        page.locator("#import-button").click()
        page.locator("#project-title").get_by_text("Big").wait_for()
        server.config["limits"]["storageUsedMb"] = 12288
        server.tasks["t0"]["status"] = "succeeded"
        playwright.expect(storage).to_contain_text("12 Go sur 138 Go", timeout=8000)

        # A deletion frees it again, without a reload.
        page.on("dialog", lambda dialog: dialog.accept())
        page.locator(".project-utilities > summary").click()
        page.locator("#delete").click()
        playwright.expect(storage).to_contain_text("0,6 Go sur 138 Go")
        assert server.errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [DESKTOP, PHONE])
def test_settings_open_from_both_entries_close_with_escape_and_return_focus(
    browser, viewport
):
    page, _ = _open(browser, viewport=viewport)
    try:
        dialog = _open_settings(page)
        playwright.expect(page.locator("#settings-limits")).to_be_visible()
        if viewport == PHONE:
            box = dialog.bounding_box()
            assert box == {"x": 0, "y": 0, "width": 390, "height": 844}
            done = page.locator("#settings-done").bounding_box()
            assert done["width"] >= 390 - 2 * 16 - 1
        for section, pane in [
            ("Documentation IA", "#settings-ai"),
            ("Sauvegarde et récupération", "#settings-backup"),
            ("Disponibilité des fonctions", "#settings-health"),
            ("Limites par opération", "#settings-limits"),
        ]:
            dialog.locator(f"[data-settings][aria-controls='{pane[1:]}']").click()
            playwright.expect(page.locator(pane)).to_be_visible()
            playwright.expect(page.locator(pane)).to_contain_text(section)
        page.keyboard.press("Escape")
        playwright.expect(dialog).to_be_hidden()
        assert page.evaluate("document.activeElement.id") == "settings-open"

        _open_limits_line(page)
        page.locator("#limits-edit").click()
        playwright.expect(dialog).to_be_visible()
        playwright.expect(page.locator("#settings-limits")).to_be_visible()
        assert (
            dialog.locator("[aria-controls='settings-limits']").get_attribute(
                "aria-current"
            )
            == "true"
        )
        page.keyboard.press("Escape")
        playwright.expect(dialog).to_be_hidden()
        assert page.evaluate("document.activeElement.id") == "limits-edit"
    finally:
        page.close()


def test_every_settings_control_shows_a_focus_ring(browser):
    page, _ = _open(browser)
    try:
        _open_settings(page)
        # Read the focused style itself, not a frame of its transition.
        page.add_style_tag(content="*{transition:none!important}")
        unringed = page.evaluate("""async () => {
          const missing = [];
          const dialog = document.querySelector('#settings');
          for (const section of dialog.querySelectorAll('[data-settings]')) {
            section.click();
            const controls = [...dialog.querySelectorAll('button,input,select')]
              .filter(e => e.checkVisibility() && !e.disabled && e.tabIndex >= 0);
            for (const control of controls) {
              control.blur();
              const resting = getComputedStyle(control).boxShadow;
              control.focus({focusVisible: true});
              const style = getComputedStyle(control);
              if (style.outlineStyle === 'none' && style.boxShadow === resting)
                missing.push(control.id || control.textContent.trim());
            }
          }
          return missing;
        }""")
        assert unringed == []
    finally:
        page.close()


def test_backup_restores_through_the_styled_picker(browser, tmp_path):
    page, server = _open(browser)
    backup = tmp_path / "budget-backup.zip"
    backup.write_bytes(b"PK\x05\x06" + bytes(18))
    restored = []

    def restore(route):
        restored.append(route.request.method)
        server.projects.insert(0, {"id": "p3", "name": "Restored", "revision": 1})
        server.json(route, {"project": server.projects[0]})

    page.route("**/api/projects/restore", restore)
    try:
        _open_settings(page)
        page.locator("[aria-controls='settings-backup']").click()
        pane = page.locator("#settings-backup")
        playwright.expect(pane.locator("#restore-file-name")).to_have_text(
            "Aucun fichier choisi"
        )
        # The native picker stays out of sight; the styled button opens it.
        native = pane.locator("input[type=file]")
        assert native.evaluate("e => getComputedStyle(e).opacity") == "0"
        with page.expect_file_chooser() as chooser:
            pane.get_by_role("button", name="+ Choisir un fichier .zip").click()
        chooser.value.set_files(backup)
        playwright.expect(pane.locator("#restore-file-name")).to_have_text(
            "budget-backup.zip"
        )
        pane.get_by_role("button", name="Restaurer l’espace").click()
        page.locator("#project-title").get_by_text("Restored").wait_for()
        assert restored == ["POST"]
        playwright.expect(page.locator("#settings")).to_be_hidden()
        assert server.errors == []
    finally:
        page.close()


def test_documentation_languages_are_named(browser):
    page, _ = _open(browser)
    try:
        _open_settings(page)
        page.locator("[aria-controls='settings-ai']").click()
        options = page.locator("#ai-language option").all_inner_texts()
        assert options == ["English", "Français", "Deutsch"]
        assert page.locator("#ai-language").input_value() == "fr"
    finally:
        page.close()


def test_lines_built_from_numbers_follow_a_language_switch(browser):
    page, server = _open(browser, lang="en")
    try:
        _open_limits_line(page)
        playwright.expect(page.locator(".limits-line")).to_contain_text(
            "Limits: 25 min · 45 s · 4.5 GB"
        )
        playwright.expect(page.locator("#connection")).to_have_text("Server connected")
        page.locator("#ui-language").select_option("fr")
        playwright.expect(page.locator(".limits-line")).to_contain_text(
            "Limites : 25 min · 45 s · 4,5 Go"
        )
        # At once: the header must not wait for the next two-second poll to
        # leave the previous language.
        playwright.expect(page.locator("#connection")).to_have_text(
            "Serveur connecté", timeout=250
        )
        playwright.expect(page.locator("#storage")).to_contain_text("1,9 Go sur 138 Go")
        _open_settings(page)
        page.locator("[aria-controls='settings-ai']").click()
        playwright.expect(page.locator("#ai-service")).to_have_text(
            "Service indisponible · Documentation : — · Vision : —"
        )
        playwright.expect(page.locator("#settings-ai-state")).to_have_text(
            "Service indisponible"
        )
        playwright.expect(page.locator("#restore-file-name")).to_have_text(
            "Aucun fichier choisi"
        )
        assert server.errors == []
    finally:
        page.close()


def _visible_text(page, selector):
    return page.locator(selector).inner_text()


@pytest.mark.parametrize("viewport", [DESKTOP, PHONE])
def test_english_sidebar_and_settings_show_no_french(browser, viewport):
    page, server = _open(browser, viewport=viewport, lang="en")
    try:
        _open_limits_line(page)
        texts = {"sidebar": _visible_text(page, "#workspace-sidebar")}
        page.locator("#settings-open").click()
        dialog = page.get_by_role("dialog", name="Settings", exact=True)
        playwright.expect(dialog).to_be_visible()
        page.locator("#limit-memory-gb").fill("99")
        page.locator("#limit-memory-gb").press("Tab")
        for pane in (
            "settings-limits",
            "settings-ai",
            "settings-backup",
            "settings-health",
        ):
            dialog.locator(f"[aria-controls='{pane}']").click()
            if pane == "settings-health":
                dialog.locator("#health-check").click()
                dialog.locator("#health-status p").wait_for()
            texts[pane] = _visible_text(page, "#settings")
        # Language names are written in their own language on purpose.
        found = {
            name: french_words(text.replace("Français", ""))
            for name, text in texts.items()
        }
        assert found == {name: [] for name in texts}, texts
        assert "Limits: 25 min · 45 s · 4.5 GB" in texts["sidebar"]
        assert "1.9 GB of 138 GB" in texts["sidebar"]
        assert (
            "“Memory per operation”: choose a value from 0.5 to 9 GB"
            in (texts["settings-limits"])
        )
        assert (
            "Server ceiling: 9 GB (option --memory-mb 9216)" in texts["settings-limits"]
        )
        assert server.errors == []
    finally:
        page.close()
