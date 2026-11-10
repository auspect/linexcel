"""Owner-local question tasks, one-question remote consent and bounded retention."""

from __future__ import annotations

import hashlib
import importlib.util
import ipaddress
import re
import secrets
import shutil
import sqlite3
import threading
import time
import uuid
from dataclasses import replace
from urllib.parse import urlsplit

from linexcel.ask import QuestionRun, encoded
from linexcel.ask_provider import OpenAIToolProvider, QuestionConfigurationChangedError
from linexcel.ask_tools import QuestionTools
from linexcel.catalog import revision
from linexcel.lazy_store import WorkbookIndex
from linexcel.question_sources import SourceSelectionError, select_question_sources

CATEGORIES = ["question", "values", "formulas", "labels", "sheet and workbook names"]


def local_endpoint(endpoint):
    host = urlsplit(endpoint).hostname
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host or "").is_loopback
    except ValueError:
        return False


class AskService:
    def __init__(self, store):
        from linexcel.observability import QuestionTracing

        self.tracing = QuestionTracing()
        self.store = store
        self.prepared = {}
        self.provider_factory = OpenAIToolProvider
        self._recover()

    def directory(self, owner, run_id):
        if not re.fullmatch(r"[a-f0-9]{32}", run_id):
            raise KeyError("Question run not found")
        return self.store.root / owner / "ask" / run_id

    def config(self):
        config = self.store.config
        available = config.ai_enabled and importlib.util.find_spec("openai") is not None
        return {
            "available": available,
            "endpoint": config.ai_base_url if config.ai_enabled else None,
            "model": config.ai_model if config.ai_enabled else None,
            "local": local_endpoint(config.ai_base_url),
            "explanation": None
            if available
            else (
                "Aucun modèle configuré pour les questions. "
                "La recherche reste disponible."
            ),
            "retentionDays": config.ask_limits.retention_days,
            "retainedRuns": config.ask_limits.retained_runs,
        }

    def _recover(self):
        from linexcel.web import _read, _write

        for path in self.store.root.glob("*/ask/*/task.json"):
            try:
                task = _read(path)
                if (
                    task["id"] != path.parent.name
                    or task["owner"] != path.parent.parent.parent.name
                    or task["operation"] != "ask"
                ):
                    continue
                if task["status"] in {"running", "queued"}:
                    task.update(
                        status="failed",
                        finishedAt=time.time(),
                        error={
                            "kind": "interrupted",
                            "message": "Question interrompue au redémarrage.",
                        },
                    )
                    _write(path, task)
                    if (path.parent / "trace.json").is_file():
                        trace = _read(path.parent / "trace.json")
                        trace.update(
                            state="terminal", verdict="failed", reason="interrupted"
                        )
                        trace["steps"].append(
                            {
                                "tool": "finish",
                                "arguments": {},
                                "summary": {
                                    "searchCompleted": False,
                                    "truncated": True,
                                    "doesNotProveAbsence": True,
                                },
                                "verdict": "failed",
                                "durationMs": 0,
                            }
                        )
                        _write(path.parent / "trace.json", trace)
                self.store.tasks[task["id"]] = task
            except (OSError, ValueError, KeyError, TypeError):
                continue

    def cleanup(self, owner):
        from linexcel.web import _read

        now = time.time()
        limits = self.store.config.ask_limits
        paths = []
        for path in (self.store.root / owner / "ask").glob("*/task.json"):
            try:
                task = _read(path)
            except (OSError, ValueError):
                continue
            if task.get("status") not in {"queued", "running"}:
                paths.append((task.get("createdAt", 0), path))
        paths.sort(reverse=True)
        for index, (created, path) in enumerate(paths):
            if (
                index >= limits.retained_runs
                or created < now - limits.retention_days * 86400
            ):
                shutil.rmtree(path.parent)
                self.store.tasks.pop(path.parent.name, None)
                self.store.events.pop(path.parent.name, None)
        self.prepared = {k: v for k, v in self.prepared.items() if v["expires"] > now}

    def prepare(self, owner, body):
        from linexcel.web import APIError

        if (
            set(body) - {"question", "project", "revision", "scope"}
            or not isinstance(body.get("question"), str)
            or not 1 <= len(body["question"].strip()) <= 2000
        ):
            raise APIError(400, "Question ou portée invalide")
        if not self.config()["available"]:
            raise APIError(409, self.config()["explanation"])
        for key in ("project", "revision", "scope"):
            if key in body and (not isinstance(body[key], str) or len(body[key]) > 128):
                raise APIError(400, "Portée invalide")
        with self.store.library.mutation(owner), self.store.lock:
            self.cleanup(owner)
            sources = self.store.library.sources(owner, body.get("scope") or None)
            # Resolve names against inventory first. An explicitly selected
            # workbook never requires opening another workbook's source index.
            try:
                inventory, _ = select_question_sources(
                    [
                        {
                            "projectId": item.project_id,
                            "id": item.project_id,
                            "filename": item.filename,
                        }
                        for item in sources
                    ],
                    body["question"],
                    project=body.get("project")
                    or (body.get("revision") or "").split(":", 1)[0]
                    or None,
                )
            except SourceSelectionError as exc:
                raise APIError(409, str(exc)) from None
            selected_ids = {item["projectId"] for item in inventory}
            sources = [item for item in sources if item.project_id in selected_ids]
            revisions = []
            for source in sources:
                with WorkbookIndex(source.path) as index:
                    sha = index.meta.get("sha256")
                    if not isinstance(sha, str) or not re.fullmatch(
                        r"[a-f0-9]{64}", sha
                    ):
                        continue
                    rev = revision(source.project_id, sha)
                    revisions.append(
                        {
                            **rev,
                            "filename": source.filename,
                            "sheets": index.meta.get("sheets", []),
                        }
                    )
            try:
                revisions, source_selection = select_question_sources(
                    revisions,
                    body["question"],
                    project=body.get("project"),
                    revision_id=body.get("revision"),
                )
            except SourceSelectionError as exc:
                raise APIError(409, str(exc)) from None
            if (
                len(revisions) > 32
                or len(encoded(revisions).encode("utf-8")) > 64 * 1024
            ):
                raise APIError(
                    413, "Portée trop grande ; choisissez moins de classeurs"
                )
            token = secrets.token_hex(32)
            prepared = {
                "owner": owner,
                "question": body["question"].strip(),
                "scope": {
                    "attachment": body.get("scope") or None,
                    "revisions": revisions,
                    "sourceSelection": source_selection,
                },
                "expires": time.time() + 600,
                "endpoint": self.store.config.ai_base_url,
                "model": self.store.config.ai_model,
                "local": local_endpoint(self.store.config.ai_base_url),
            }
            # At most 32 outstanding preparations per owner; no stale consent reuse.
            owner_tokens = [k for k, v in self.prepared.items() if v["owner"] == owner]
            for old in owner_tokens[:-31]:
                self.prepared.pop(old, None)
            self.prepared[token] = prepared
            return {
                "token": token,
                "question": prepared["question"],
                "endpoint": prepared["endpoint"],
                "model": prepared["model"],
                "local": prepared["local"],
                "revisions": revisions,
                "sourceSelection": source_selection,
                "categories": CATEGORIES,
                "expiresAt": prepared["expires"],
            }

    def submit(self, owner, body):
        from linexcel.web import APIError, _write

        if (
            set(body) - {"token", "remoteConsent"}
            or not isinstance(body.get("token"), str)
            or type(body.get("remoteConsent", False)) is not bool
        ):
            raise APIError(400, "Autorisation de question invalide")
        with self.store.lock:
            prepared = self.prepared.get(body["token"])
            if (
                not prepared
                or prepared["owner"] != owner
                or prepared["expires"] <= time.time()
            ):
                raise APIError(409, "Préparez à nouveau cette question")
            if not prepared["local"] and body.get("remoteConsent") is not True:
                raise APIError(409, "Autorisation distante requise pour cette question")
            if (
                prepared["endpoint"] != self.store.config.ai_base_url
                or prepared["model"] != self.store.config.ai_model
                or not self.config()["available"]
            ):
                raise APIError(409, "La configuration du modèle a changé")
            self.store._has_capacity()
            run_id = uuid.uuid4().hex
            root = self.directory(owner, run_id)
            root.mkdir(parents=True)
            task = {
                "id": run_id,
                "owner": owner,
                "projectId": None,
                "revision": "question",
                "operation": "ask",
                "nodeId": None,
                "status": "queued",
                "createdAt": time.time(),
                "error": None,
                "progress": {"phase": "queued", "fraction": 0},
                "budget": self.store.budget(
                    {
                        "seconds": self.store.config.ask_limits.seconds,
                        "memoryMb": min(256, self.store.config.total_memory_mb),
                    }
                ),
                "options": {},
                "result": None,
            }
            self.store.tasks[run_id] = task
            self.store.events[run_id] = threading.Event()
            _write(
                root / "input.json",
                {
                    **prepared,
                    "remoteConsent": body.get("remoteConsent", False),
                    "categories": CATEGORIES,
                },
            )
            _write(root / "task.json", task)
            _write(
                root / "trace.json",
                {
                    "schemaVersion": "linexcel.ask.v1",
                    "state": "queued",
                    "scope": prepared["scope"],
                    "steps": [],
                    "budgetUse": {"steps": 0, "tokens": 0, "resultBytes": 0},
                    "verdict": None,
                },
            )
            self.prepared.pop(body["token"])
            self.store.executor.submit(self.store._run, task, root)
            return {"task": self.store.public(task, include_result=False)}

    def run(self, task, event):
        from linexcel.web import _read, _write

        root = self.directory(task["owner"], task["id"])
        prepared = _read(root / "input.json")
        with self.store.lock:
            approved_config = replace(
                self.store.config,
                ai_base_url=prepared["endpoint"],
                ai_model=prepared["model"],
            )
        limits = replace(
            self.store.config.ask_limits,
            tokens=min(self.store.config.ask_limits.tokens, task["budget"]["tokens"]),
        )
        tools = QuestionTools(
            self.store,
            task["owner"],
            prepared["scope"],
            root,
            time.monotonic() + limits.seconds,
        )

        def publish(trace):
            with self.store.lock:
                _write(root / "trace.json", trace)
                task["progress"] = {
                    "phase": trace["state"],
                    "fraction": None,
                    "steps": len(trace["steps"]),
                    "budgetUse": dict(trace["budgetUse"]),
                }

        provider = None

        def before_request():
            with self.store.lock:
                current = self.store.config
                if (
                    not current.ai_enabled
                    or current.ai_base_url != prepared["endpoint"]
                    or current.ai_model != prepared["model"]
                    or (not prepared["local"] and prepared["remoteConsent"] is not True)
                ):
                    raise QuestionConfigurationChangedError()

        try:
            # Construction is also outside TaskStore.lock. Failures still have
            # an application-owned terminal trace.
            try:
                provider = self.provider_factory(approved_config)
            except Exception:
                from linexcel.aidoc import TokenUsage
                from linexcel.ask_provider import ToolResponse

                class Unavailable:
                    def complete(self, *args, **kwargs):
                        return ToolResponse([], "stop", TokenUsage())

                    def close(self):
                        pass

                provider = Unavailable()
            result = QuestionRun(
                provider,
                tools,
                prepared["question"],
                limits,
                event,
                publish,
                before_request=before_request,
                observation=self.tracing.start(),
            ).run()
            result["disclosure"] = {
                k: prepared[k]
                for k in ("endpoint", "model", "local", "remoteConsent", "categories")
            }
            with self.store.lock:
                _write(root / "terminal.json", result)
            return {"result": result}
        finally:
            if provider is not None:
                provider.close()

    def result(self, owner, run_id):
        from linexcel.web import _read

        root = self.directory(owner, run_id)
        if not (root / "trace.json").is_file():
            return {"status": "running", "answer": None, "trace": None}
        with self.store.lock:
            trace = _read(root / "trace.json")
            result = (
                _read(root / "terminal.json")
                if (root / "terminal.json").is_file()
                else {
                    "status": "failed" if trace["state"] == "terminal" else "running",
                    "answer": None,
                    "trace": trace,
                }
            )
        # Access rules match source evidence, including replacement and attachment
        # expiry. Persisted run logs never extend source authorization.
        try:
            tools = QuestionTools(
                self.store, owner, trace["scope"], root, time.monotonic() + 5
            )
            with tools.authorize() as sources:
                answer = result.get("answer")
                if result.get("status") == "answer_proposed" and (
                    not isinstance(answer, dict)
                    or answer.get("evidenceState")
                    != tools.current_evidence_state(sources, answer)
                ):
                    return {
                        "status": "failed",
                        "answer": None,
                        "trace": {
                            **trace,
                            "state": "terminal",
                            "verdict": "failed",
                            "reason": "lineage_state_changed",
                            "steps": [
                                {
                                    "tool": "finish",
                                    "verdict": "failed",
                                    "summary": {
                                        "searchCompleted": False,
                                        "truncated": True,
                                        "doesNotProveAbsence": True,
                                    },
                                }
                            ],
                        },
                    }
        except (KeyError, ValueError, OSError, sqlite3.Error, TimeoutError):
            return {
                "status": "failed",
                "answer": None,
                "trace": {
                    "schemaVersion": trace["schemaVersion"],
                    "state": "terminal",
                    "verdict": "failed",
                    "reason": "scope_or_source_changed",
                    "steps": [
                        {
                            "tool": "finish",
                            "summary": {
                                "searchCompleted": False,
                                "truncated": True,
                                "doesNotProveAbsence": True,
                            },
                            "verdict": "failed",
                        }
                    ],
                },
            }
        with self.store.lock:
            if (root / "feedback.json").is_file():
                result["feedback"] = _read(root / "feedback.json")
        return result

    def finalize(self, task):
        """Even cancellation before scheduling has a persisted terminal trace."""
        from linexcel.web import _read, _write

        if task["status"] == "succeeded":
            return
        root = self.directory(task["owner"], task["id"])
        trace = _read(root / "trace.json")
        reason = "cancelled" if task["status"] == "cancelled" else "interrupted"
        trace.update(state="terminal", verdict="failed", reason=reason)
        if not trace["steps"] or trace["steps"][-1]["tool"] != "finish":
            trace["steps"].append(
                {
                    "tool": "finish",
                    "arguments": {},
                    "summary": {
                        "searchCompleted": False,
                        "truncated": True,
                        "doesNotProveAbsence": True,
                    },
                    "verdict": "failed",
                    "durationMs": 0,
                }
            )
        _write(root / "trace.json", trace)
        _write(
            root / "terminal.json", {"status": "failed", "answer": None, "trace": trace}
        )

    def get(self, owner, run_id):
        from linexcel.web import APIError

        with self.store.lock:
            self.cleanup(owner)
            task = self.store.tasks.get(run_id)
            if task is None or task["owner"] != owner or task["operation"] != "ask":
                raise APIError(404, "Question introuvable")
            public = self.store.public(task, include_result=False)
        # Authorization acquires the owner mutation gate before the store lock.
        return {"task": public, "run": self.result(owner, run_id)}

    def feedback(self, owner, run_id, body):
        from linexcel.web import APIError, _read, _write

        if (
            set(body) != {"decision"}
            or not isinstance(body["decision"], str)
            or body["decision"] not in {"confirm", "reject"}
        ):
            raise APIError(400, "Décision invalide")
        run = self.get(owner, run_id)["run"]
        if run["status"] != "answer_proposed":
            raise APIError(409, "Aucune proposition disponible")
        with self.store.lock:
            root = self.directory(owner, run_id)
            path = root / "feedback.json"
            history = _read(path) if path.is_file() else []
            if len(history) >= 20:
                raise APIError(409, "Limite de décisions atteinte")
            history.append(
                {
                    "decision": body["decision"],
                    "at": time.time(),
                    "questionSha256": hashlib.sha256(
                        (_read(root / "input.json"))["question"].encode()
                    ).hexdigest(),
                }
            )
            _write(path, history)
        return {"feedback": history}
