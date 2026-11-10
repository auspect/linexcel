"""A source specified by the user becomes a server-owned question boundary."""

import hashlib
import json
import sqlite3
from pathlib import Path

import pytest
from test_ask import FakeProvider, call
from test_web import request, wait, workbook
from test_web_ask import app as app

from linexcel.web_ask import WorkbookIndex


def add(app, name):
    data = workbook()
    data["name"] = name
    status, raw, _ = request(app, "POST", "/api/projects", {"workbook": data})
    assert status == 200
    project = json.loads(raw)["project"]
    owner = hashlib.sha256(b"alice").hexdigest()
    task = app.store.submit(owner, project["id"], "import", None, {})
    assert wait(app.store, task["id"])["status"] == "succeeded"
    return project


def test_named_source_freezes_scope_and_opens_only_selected_index(app, monkeypatch):
    chosen = add(app, "chosen.xlsx")
    other = add(app, "other.xlsx")
    opened = []

    def tracked(path, *args, **kwargs):
        opened.append(Path(path))
        return WorkbookIndex(path, *args, **kwargs)

    monkeypatch.setattr("linexcel.web_ask.WorkbookIndex", tracked)
    code, raw, _ = request(
        app,
        "POST",
        "/api/library/ask/prepare",
        {
            "question": "Dans chosen.xlsx uniquement, quelle est la valeur ?",
        },
    )
    assert code == 200, raw
    prepared = json.loads(raw)
    assert [item["projectId"] for item in prepared["revisions"]] == [chosen["id"]]
    assert prepared["sourceSelection"] == {
        "kind": "question_filename",
        "filenames": ["chosen.xlsx"],
    }
    assert len(opened) == 1 and opened[0].parent.name == chosen["id"]
    snapshot = app.store.ask.prepared[prepared["token"]]["scope"]
    assert other["id"] not in json.dumps(snapshot)

    # A model attempt to inspect the excluded workbook is rejected, even though
    # the workbook belongs to the same owner and exists in the library.
    app.store.ask.provider_factory = lambda _: FakeProvider(
        [
            [
                call(
                    "read_range",
                    {"project": other["id"], "sheet": "Sheet", "range": "A1"},
                )
            ]
        ]
    )
    code, raw, _ = request(
        app, "POST", "/api/library/ask", {"token": prepared["token"]}
    )
    assert code == 200, raw
    task = json.loads(raw)["task"]
    wait(app.store, task["id"])
    code, raw, _ = request(app, "GET", f"/api/library/ask/{task['id']}")
    result = json.loads(raw)["run"]
    assert result["status"] == "failed" and result["answer"] is None
    assert result["trace"]["scope"]["sourceSelection"] == prepared["sourceSelection"]


@pytest.mark.parametrize(
    "question, explicit",
    [
        ("Dans missing.xlsx", False),
        ('Dans "Missing same.xlsx"', False),
        ("Dans same.xlsx", False),
        ("Dans other.xlsx", True),
    ],
)
def test_unresolved_source_never_prepares_token_or_calls_provider(
    app, question, explicit
):
    chosen = add(app, "same.xlsx")
    add(app, "same.xlsx")
    add(app, "other.xlsx")
    constructed = []
    app.store.ask.provider_factory = lambda config: constructed.append(config)
    body = {"question": question}
    if explicit:
        body["project"] = chosen["id"]
    code, _, _ = request(app, "POST", "/api/library/ask/prepare", body)
    assert code == 409
    assert not app.store.ask.prepared and not constructed
    assert not any(task["operation"] == "ask" for task in app.store.tasks.values())


def test_explicit_project_disambiguates_duplicate_filename(app):
    chosen = add(app, "same.xlsx")
    add(app, "same.xlsx")
    code, raw, _ = request(
        app,
        "POST",
        "/api/library/ask/prepare",
        {
            "question": "Dans same.xlsx, quelle valeur ?",
            "project": chosen["id"],
        },
    )
    assert code == 200, raw
    prepared = json.loads(raw)
    assert prepared["sourceSelection"]["kind"] == "explicit"
    assert [item["projectId"] for item in prepared["revisions"]] == [chosen["id"]]


def test_attached_reference_name_does_not_block_the_question(app):
    report = add(app, "rapport.xlsx")
    add(app, "autre.xlsx")
    owner = hashlib.sha256(b"alice").hexdigest()
    references = app.store.project_path(owner, report["id"]) / "references"
    references.mkdir(exist_ok=True)
    (references / "budget.xlsx").write_bytes(b"attached reference")
    code, raw, _ = request(
        app,
        "POST",
        "/api/library/ask/prepare",
        {"question": "Dans rapport.xlsx, d’où vient B5 via budget.xlsx ?"},
    )
    assert code == 200, raw
    prepared = json.loads(raw)
    assert [item["projectId"] for item in prepared["revisions"]] == [report["id"]]


def test_index_still_being_written_is_a_clear_conflict(app, monkeypatch):
    pending = add(app, "pending.xlsx")

    def still_building(path, *args, **kwargs):
        # What a live import produced on 2026-11-10: the source was listed
        # while its index could not be opened yet.
        if Path(path).parent.name == pending["id"]:
            raise sqlite3.OperationalError("unable to open database file")
        return WorkbookIndex(path, *args, **kwargs)

    monkeypatch.setattr("linexcel.web_ask.WorkbookIndex", still_building)
    code, raw, _ = request(
        app,
        "POST",
        "/api/library/ask/prepare",
        {"question": "Dans pending.xlsx, quelle est la valeur ?"},
    )
    assert code == 409, raw
    assert "« pending.xlsx »" in json.loads(raw)["error"]["message"]
    assert not app.store.ask.prepared
