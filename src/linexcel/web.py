"""Embeddable, deterministic lazy application and persistent task coordinator.

One coordinator owns a data directory. Host applications authenticate requests
with ``resolve_user(environ)``; the standalone loopback server uses a private
browser cookie. No request can select another user's on-disk directory.
"""

from __future__ import annotations

import base64
import hashlib
import json
import logging
import math
import os
import re
import secrets
import shutil
import signal
import subprocess
import sys
import threading
import time
import uuid
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from http.cookies import SimpleCookie
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

TERMINAL = {"succeeded", "failed", "cancelled"}
OPERATIONS = {"import", "evaluate", "capture", "document", "describe_capture"}
AI_OPERATIONS = {"document", "describe_capture"}


def _write(path: Path, value: Any) -> None:
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, allow_nan=False), "utf-8")
    temp.replace(path)


def _read(path: Path) -> Any:
    return json.loads(path.read_text("utf-8"))


@dataclass(frozen=True)
class WebConfig:
    data_dir: Path
    mount_path: str = ""
    max_workers: int = 2
    total_memory_mb: int = 2048
    max_upload_mb: int = 64
    max_storage_mb: int = 2048
    retention_days: int = 7
    max_pending_tasks: int = 32
    ai_enabled: bool = False
    ai_base_url: str = field(default="http://localhost:11434/v1", repr=False)
    ai_api_key: str | None = field(default=None, repr=False)
    ai_model: str = "qwen3.8"
    ai_vision_model: str | None = None
    ai_token_budget: int = 16000

    def __post_init__(self):
        for key in (
            "max_workers",
            "total_memory_mb",
            "max_upload_mb",
            "max_storage_mb",
            "retention_days",
            "max_pending_tasks",
            "ai_token_budget",
        ):
            value = getattr(self, key)
            if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
                raise ValueError(f"{key} must be a positive integer")
        if not isinstance(self.ai_enabled, bool):
            raise ValueError("ai_enabled must be boolean")
        if self.ai_enabled:
            endpoint = urlsplit(self.ai_base_url)
            if (
                endpoint.scheme not in {"http", "https"}
                or not endpoint.netloc
                or endpoint.username
                or endpoint.password
                or endpoint.query
                or endpoint.fragment
            ):
                raise ValueError(
                    "AI endpoint must be HTTP(S) without URL credentials, "
                    "query or fragment"
                )
            if not self.ai_model or len(self.ai_model) > 200:
                raise ValueError(
                    "An AI model name is required (at most 200 characters)"
                )
            if self.ai_vision_model is not None and (
                not self.ai_vision_model or len(self.ai_vision_model) > 200
            ):
                raise ValueError("Invalid vision model name")
        if self.mount_path and not re.fullmatch(
            r"(?:/[A-Za-z0-9_-]+)+", self.mount_path
        ):
            raise ValueError(
                "mount_path must be empty or /path without a trailing slash"
            )


class APIError(Exception):
    def __init__(self, status: int, message: str):
        self.status = status
        super().__init__(message)


class TaskStore:
    """A single process owns scheduling; task results survive browser reconnects."""

    def __init__(self, config: WebConfig):
        self.config = config
        self.root = Path(config.data_dir).resolve()
        if os.name == "nt" and not str(self.root).startswith("\\\\?\\"):
            # User/project/task IDs can exceed legacy MAX_PATH in a nested host
            # data directory. The explicit prefix works without registry edits.
            raw = str(self.root)
            self.root = Path(
                "\\\\?\\UNC\\" + raw[2:] if raw.startswith("\\\\") else "\\\\?\\" + raw
            )
        self.root.mkdir(parents=True, exist_ok=True)
        self._lock_file = (self.root / "coordinator.lock").open("a+b")
        try:
            if os.name == "nt":
                import msvcrt

                self._lock_file.seek(0)
                self._lock_file.write(b"0")
                self._lock_file.flush()
                self._lock_file.seek(0)
                msvcrt.locking(self._lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(self._lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self._lock_file.close()
            raise RuntimeError(
                "This data directory already has a coordinator"
            ) from None
        self.lock = threading.RLock()
        self.capacity = threading.Condition(self.lock)
        self.used_memory = 0
        self.closed = False
        self.events: dict[str, threading.Event] = {}
        self.tasks: dict[str, dict] = {}
        from importlib.metadata import version

        cache = hashlib.sha256(version("formualizer").encode())
        for name in (
            "lazy.py",
            "engine.py",
            "excel_compat.py",
            "values.py",
            "loader.py",
            "external.py",
            "refs.py",
            "web_worker.py",
            "web.py",
            "web_ai.py",
            "aidoc.py",
        ):
            cache.update(Path(__file__).with_name(name).read_bytes())
        for prompt in sorted(
            (Path(__file__).parent / "assets" / "prompts").glob("*/*.md")
        ):
            cache.update(prompt.read_bytes())
        self.cache_version = cache.hexdigest()
        self.executor = ThreadPoolExecutor(max_workers=config.max_workers)
        self._recover()

    def _recover(self):
        cutoff = time.time() - self.config.retention_days * 86400
        for path in self.root.glob("*/projects/*/project.json"):
            project = _read(path)
            if project["createdAt"] < cutoff:
                shutil.rmtree(path.parent)
                continue
            for task_path in path.parent.glob("tasks/*/task.json"):
                task = _read(task_path)
                if task["status"] not in TERMINAL:
                    task.update(
                        status="failed",
                        error={
                            "kind": "interrupted",
                            "message": "Serveur redémarré : opération interrompue, "
                            "à relancer.",
                        },
                    )
                    _write(task_path, task)
                self.tasks[task["id"]] = task

    def close(self):
        with self.capacity:
            self.closed = True
            for event in self.events.values():
                event.set()
            self.capacity.notify_all()
        self.executor.shutdown(wait=True)
        self._lock_file.close()

    def project_path(self, owner: str, project_id: str) -> Path:
        if not re.fullmatch(r"[a-f0-9]{32}", project_id):
            raise APIError(404, "Classeur introuvable")
        path = self.root / owner / "projects" / project_id
        if not (path / "project.json").is_file():
            raise APIError(404, "Classeur introuvable")
        return path

    def budget(self, value: Any) -> dict:
        if not isinstance(value, dict):
            raise APIError(400, "Budget invalide")
        seconds, memory = value.get("seconds", 120), value.get("memoryMb", 512)
        if (
            isinstance(seconds, bool)
            or not isinstance(seconds, (float, int))
            or not math.isfinite(seconds)
            or not 0 < seconds <= 3600
            or isinstance(memory, bool)
            or not isinstance(memory, int)
            or not 128 <= memory <= self.config.total_memory_mb
        ):
            raise APIError(
                400, "Budget : 0–3600 secondes et mémoire dans la limite serveur"
            )
        tokens = value.get("tokens", self.config.ai_token_budget)
        if (
            isinstance(tokens, bool)
            or not isinstance(tokens, int)
            or not 1 <= tokens <= self.config.ai_token_budget
        ):
            raise APIError(400, "Budget de tokens hors limite serveur")
        return {"seconds": seconds, "memoryMb": memory, "tokens": tokens}

    def _has_capacity(self):
        if self.closed:
            raise APIError(503, "Serveur en cours d’arrêt")
        pending = sum(t["status"] not in TERMINAL for t in self.tasks.values())
        if pending >= self.config.max_pending_tasks:
            raise APIError(429, "File des opérations pleine ; réessayez plus tard")
        if self._storage_bytes() >= self.config.max_storage_mb * 1048576:
            raise APIError(413, "Stockage plein ; supprimez des classeurs")

    def _storage_bytes(self) -> int:
        size = 0
        for path in self.root.rglob("*"):
            try:
                if path.is_file():
                    size += path.stat().st_size
            except FileNotFoundError:
                pass  # Another task atomically published its result.
        return size

    @staticmethod
    def _file(value: Any) -> tuple[str, bytes]:
        if not isinstance(value, dict):
            raise APIError(400, "Fichier invalide")
        name = value.get("name", "")
        if (
            not isinstance(name, str)
            or len(name) > 180
            or not re.fullmatch(r"[^/\\:\x00-\x1f]+\.(?:xlsx|xlsm)", name, re.I)
            or name.startswith(".")
            or name.rstrip(". ") != name
        ):
            raise APIError(400, "Nom de fichier .xlsx/.xlsm invalide")
        try:
            data = base64.b64decode(value["data"], validate=True)
        except (KeyError, ValueError, TypeError):
            raise APIError(400, "Fichier base64 invalide") from None
        if not data.startswith(b"PK"):
            raise APIError(400, "Le fichier doit être un classeur OOXML")
        return name, data

    def create(self, owner: str, body: dict) -> dict:
        budget = self.budget(body.get("budget", {}))
        workbook = self._file(body.get("workbook"))
        workspace_name = body.get("name") or workbook[0]
        if not isinstance(workspace_name, str) or len(workspace_name) > 200:
            raise APIError(400, "Nom de l’espace invalide")
        refs = body.get("references", [])
        if not isinstance(refs, list) or len(refs) > 32:
            raise APIError(400, "Au plus 32 classeurs de référence")
        files = [workbook, *(self._file(item) for item in refs)]
        if len({name.casefold() for name, _ in files}) != len(files):
            raise APIError(400, "Les noms des fichiers doivent être uniques")
        if sum(len(data) for _, data in files) > self.config.max_upload_mb * 1048576:
            raise APIError(413, "Import trop volumineux")
        with self.lock:
            self._has_capacity()
            used = self._storage_bytes()
            if (
                used + sum(len(d) for _, d in files)
                > self.config.max_storage_mb * 1048576
            ):
                raise APIError(413, "Stockage plein ; supprimez des classeurs")
            project_id = uuid.uuid4().hex
            path = self.root / owner / "projects" / project_id
            (path / "references").mkdir(parents=True)
            (path / "workbook.xlsx").write_bytes(workbook[1])
            for name, data in files[1:]:
                (path / "references" / name).write_bytes(data)
            revision = self._revision(path)
            project = {
                "id": project_id,
                "name": workspace_name,
                "filename": workbook[0],
                "revision": revision,
                "references": [name for name, _ in files[1:]],
                "createdAt": time.time(),
            }
            _write(path / "project.json", project)
            task = self.submit(owner, project_id, "import", None, budget)
            return {"project": project, "task": task}

    @staticmethod
    def _revision(path: Path) -> str:
        digest = hashlib.sha256()
        for file in [path / "workbook.xlsx", *sorted((path / "references").iterdir())]:
            digest.update(file.name.encode("utf-8"))
            with file.open("rb") as stream:
                while chunk := stream.read(1024 * 1024):
                    digest.update(chunk)
        return digest.hexdigest()

    def list_projects(self, owner: str) -> list:
        with self.lock:
            return sorted(
                [_read(p) for p in (self.root / owner).glob("projects/*/project.json")],
                key=lambda p: p["createdAt"],
                reverse=True,
            )

    def snapshot(self, owner: str, project_id: str) -> dict:
        with self.lock:
            path = self.project_path(owner, project_id)
            return {
                "project": _read(path / "project.json"),
                "graph": _read(path / "graph.json")
                if (path / "graph.json").exists()
                else None,
                "tasks": [
                    self.public(t)
                    for t in self.tasks.values()
                    if t["owner"] == owner and t["projectId"] == project_id
                ],
            }

    @staticmethod
    def public(task: dict) -> dict:
        return {k: v for k, v in task.items() if k != "owner"}

    def get_task(self, owner: str, task_id: str) -> dict:
        with self.lock:
            task = self.tasks.get(task_id)
            if task is None or task["owner"] != owner:
                raise APIError(404, "Opération introuvable")
            return task.copy()

    def submit(
        self,
        owner: str,
        project_id: str,
        operation: str,
        node_id: str | None,
        budget: dict,
        force: bool = False,
        options: dict | None = None,
    ) -> dict:
        if operation in AI_OPERATIONS and isinstance(budget, dict):
            budget = {"seconds": 600, **budget}
        budget = self.budget(budget)
        options = {} if options is None else options
        if not isinstance(options, dict):
            raise APIError(400, "Options invalides")
        allowed = (
            {"language", "captureTaskId", "captureIndex"}
            if operation == "describe_capture"
            else ({"language"} if operation == "document" else set())
        )
        if set(options) - allowed:
            raise APIError(
                400,
                "Option non autorisée ; fournisseur et modèle "
                "sont configurés côté serveur",
            )
        if operation in AI_OPERATIONS:
            from linexcel.i18n import LANGUAGES

            if not self.config.ai_enabled:
                raise APIError(409, "Aucun fournisseur IA activé côté serveur")
            language = options.get("language", "fr")
            if not isinstance(language, str) or language not in LANGUAGES:
                raise APIError(400, "Langue non prise en charge")
            options = {**options, "language": language}
        if not isinstance(force, bool):
            raise APIError(400, "force doit être booléen")
        if operation not in OPERATIONS:
            raise APIError(400, "Opération indisponible")
        with self.lock:
            path = self.project_path(owner, project_id)
            project = _read(path / "project.json")
            if self._revision(path) != project["revision"]:
                raise APIError(
                    409, "Sources modifiées sur disque ; importez une nouvelle version"
                )
            pattern_id = None
            if operation == "evaluate" or (
                operation == "document" and node_id is not None
            ):
                if not isinstance(node_id, str) or not (path / "graph.json").exists():
                    raise APIError(409, "Importez le graphe avant le recalcul")
                node = next(
                    (
                        n
                        for n in _read(path / "graph.json")["nodes"]
                        if n["id"] == node_id
                    ),
                    None,
                )
                if node is None:
                    raise APIError(404, "Nœud introuvable")
                pattern_id = node.get("patternId")
            ai_snapshot = None
            ai_cache_key = None
            if operation in AI_OPERATIONS:
                from linexcel.web_ai import capture_evidence, digest, evidence

                if operation == "document":
                    if not (path / "graph.json").is_file():
                        raise APIError(409, "Importez le graphe avant la documentation")
                    completed = [
                        t
                        for t in self.tasks.values()
                        if t["owner"] == owner
                        and t["projectId"] == project_id
                        and t["revision"] == project["revision"]
                        and t.get("cacheVersion") == self.cache_version
                        and t["operation"] == "evaluate"
                        and (node_id is None or t["nodeId"] == node_id)
                        and t["status"] == "succeeded"
                        and t.get("result")
                    ]
                    ai_snapshot = evidence(
                        _read(path / "graph.json"), node_id, completed
                    )
                else:
                    if node_id is not None:
                        raise APIError(
                            400, "La description d’image ne prend pas de nodeId"
                        )
                    capture_id, index = (
                        options.get("captureTaskId"),
                        options.get("captureIndex"),
                    )
                    if (
                        not isinstance(capture_id, str)
                        or isinstance(index, bool)
                        or not isinstance(index, int)
                    ):
                        raise APIError(400, "Capture et index requis")
                    captured = self.get_task(owner, capture_id)
                    if (
                        captured["projectId"] != project_id
                        or captured["revision"] != project["revision"]
                        or captured["operation"] != "capture"
                        or captured["status"] != "succeeded"
                    ):
                        raise APIError(
                            409,
                            "Capture réussie requise pour cette version du classeur",
                        )
                    shots = (captured.get("result") or {}).get("screenshots", [])
                    if not 0 <= index < len(shots):
                        raise APIError(400, "Index de capture invalide")
                    ai_snapshot = capture_evidence(captured, index)
                ai_cache_key = digest(
                    {
                        "evidence": ai_snapshot,
                        "options": options,
                        "endpoint": self.config.ai_base_url,
                        "model": self.config.ai_model,
                        "visionModel": self.config.ai_vision_model
                        or self.config.ai_model,
                        "credentialDigest": hashlib.sha256(
                            (self.config.ai_api_key or "").encode()
                        ).hexdigest(),
                    }
                )
            # Immutable content revisions make cached results safe to reuse.
            # Recovery enumerates UUID directories, whose filesystem order is
            # unrelated to request order. Reuse the most recent matching task.
            for task in sorted(
                self.tasks.values(),
                key=lambda item: (item["createdAt"], item["id"]),
                reverse=True,
            ):
                if (
                    task["owner"] == owner
                    and task["projectId"] == project_id
                    and task["operation"] == operation
                    and task["nodeId"] == node_id
                    and task["revision"] == project["revision"]
                    and task.get("cacheVersion") == self.cache_version
                    and task.get("aiCacheKey") == ai_cache_key
                    and task.get("options", {}) == options
                    and task["status"] in {"queued", "running", "succeeded"}
                    and not (force and task["status"] == "succeeded")
                ):
                    return self.public(task)
            self._has_capacity()
            task_id = uuid.uuid4().hex
            task = {
                "id": task_id,
                "owner": owner,
                "projectId": project_id,
                "revision": project["revision"],
                "cacheVersion": self.cache_version,
                "aiCacheKey": ai_cache_key,
                "options": options,
                "operation": operation,
                "nodeId": node_id,
                "patternId": pattern_id,
                "status": "queued",
                "createdAt": time.time(),
                "progress": {"phase": "queued", "fraction": 0},
                "budget": budget,
                "error": None,
                "result": None,
            }
            task_dir = path / "tasks" / task_id
            task_dir.mkdir(parents=True)
            self.tasks[task_id] = task
            self.events[task_id] = threading.Event()
            _write(task_dir / "task.json", task)
            if ai_snapshot is not None:
                _write(task_dir / "ai_input.json", ai_snapshot)
            self.executor.submit(self._run, task, task_dir)
            return self.public(task)

    def cancel(self, owner: str, task_id: str) -> dict:
        with self.capacity:
            task = self.get_task(owner, task_id)
            if task["status"] not in TERMINAL:
                self.events[task_id].set()
                self.capacity.notify_all()
            return self.public(task)

    def delete(self, owner: str, project_id: str):
        with self.lock:
            path = self.project_path(owner, project_id)
            tasks = [
                t
                for t in self.tasks.values()
                if t["owner"] == owner and t["projectId"] == project_id
            ]
            active = [t for t in tasks if t["status"] not in TERMINAL]
            for task in active:
                self.cancel(owner, task["id"])
            if active:
                raise APIError(409, "Annulation demandée ; réessayez après leur arrêt")
            # path uses server-generated IDs, verified against this owner's root.
            path.resolve().relative_to(self.root / owner / "projects")
            shutil.rmtree(path)
            for task in tasks:
                self.tasks.pop(task["id"], None)
                self.events.pop(task["id"], None)

    def _run(self, task: dict, root: Path):
        event = self.events[task["id"]]
        memory = task["budget"]["memoryMb"]
        reserved = False
        try:
            with self.capacity:
                while self.used_memory + memory > self.config.total_memory_mb:
                    if event.is_set():
                        break
                    self.capacity.wait(0.1)
                if event.is_set():
                    task["status"] = "cancelled"
                    return
                self.used_memory += memory
                reserved = True
                task.update(status="running", startedAt=time.time())
                task["progress"] = {"phase": task["operation"], "fraction": None}
                _write(root / "task.json", task)
            result = self._worker(task, root, event)
            with self.lock:
                if event.is_set():
                    task["status"] = "cancelled"
                elif result.get("error"):
                    task.update(status="failed", error=result["error"])
                else:
                    payload = result["result"]
                    if task["operation"] == "import":
                        _write(root.parent.parent / "graph.json", payload)
                        payload = {"nodeCount": len(payload["nodes"])}
                    task.update(status="succeeded", result=payload)
        except Exception as exc:
            with self.lock:
                task.update(
                    status="failed",
                    error={"kind": type(exc).__name__, "message": str(exc)},
                )
        finally:
            with self.capacity:
                try:
                    # Server-generated task paths, never uploaded filenames.
                    root.resolve().relative_to(self.root)
                    for scratch in ("result.json", "request.json", "ai_input.json"):
                        (root / scratch).unlink(missing_ok=True)
                    if (root / "captures").exists():
                        shutil.rmtree(root / "captures")
                    log = root / "stderr.txt"
                    if task["operation"] in AI_OPERATIONS:
                        log.unlink(missing_ok=True)
                    if log.exists() and log.stat().st_size > 16384:
                        with log.open("rb") as stream:
                            head = stream.read(16384)
                        log.write_bytes(head)
                    task["finishedAt"] = time.time()
                    task["progress"] = {"phase": task["status"], "fraction": 1}
                    _write(root / "task.json", task)
                except OSError as exc:
                    task.update(
                        status="failed",
                        error={
                            "kind": "persistence_error",
                            "message": str(exc),
                        },
                        progress={"phase": "failed", "fraction": 1},
                    )
                    logging.getLogger(__name__).exception("Task persistence failed")
                finally:
                    if reserved:
                        self.used_memory -= memory
                    self.capacity.notify_all()

    def _worker(self, task: dict, root: Path, event: threading.Event) -> dict:
        request = {
            "operation": task["operation"],
            "nodeId": task["nodeId"],
            "memoryMb": task["budget"]["memoryMb"],
            "options": task.get("options", {}),
            "tokens": task["budget"].get("tokens", self.config.ai_token_budget),
        }
        _write(root / "request.json", request)
        child_env = os.environ.copy()
        child_env.pop("LINEXCEL_WEB_AI_CONFIG", None)
        if task["operation"] in AI_OPERATIONS:
            child_env["LINEXCEL_WEB_AI_CONFIG"] = json.dumps(
                {
                    "baseUrl": self.config.ai_base_url,
                    "apiKey": self.config.ai_api_key,
                    "model": self.config.ai_model,
                    "visionModel": self.config.ai_vision_model or self.config.ai_model,
                }
            )
            child_env["LINEXCEL_AI_TIMEOUT_SECONDS"] = str(task["budget"]["seconds"])
        command = [
            sys.executable,
            "-X",
            "faulthandler",
            "-m",
            "linexcel.web_worker",
            str(root),
        ]
        if os.name != "nt":
            bootstrap = (
                "import resource,runpy,sys;resource.setrlimit(resource.RLIMIT_AS,"
                f"({request['memoryMb'] * 1048576},)*2);"
                "sys.argv=['linexcel.web_worker',sys.argv[1]];"
                "runpy.run_module('linexcel.web_worker',run_name='__main__')"
            )
            command = [sys.executable, "-X", "faulthandler", "-c", bootstrap, str(root)]
        started = time.monotonic()
        job = None
        reason = None
        with (root / "stderr.txt").open("wb") as log:
            process = subprocess.Popen(
                command,
                env=child_env,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=log,
                start_new_session=os.name != "nt",
                creationflags=(subprocess.CREATE_NO_WINDOW | 0x4)
                if os.name == "nt"
                else 0,
            )
            try:
                if os.name == "nt":
                    from linexcel._windows_job import WorkerJob

                    job = WorkerJob(process, request["memoryMb"])
                    job.resume(process)
                while process.poll() is None:
                    if event.wait(0.05):
                        reason = "cancelled"
                        break
                    if time.monotonic() - started >= task["budget"]["seconds"]:
                        reason = "timeout"
                        break
                    output_size = sum(
                        p.stat().st_size for p in root.rglob("*") if p.is_file()
                    )
                    if output_size > self.config.max_upload_mb * 4 * 1048576:
                        reason = "output_limit"
                        break
                    if self._storage_bytes() > self.config.max_storage_mb * 1048576:
                        reason = "storage_limit"
                        break
            finally:
                if job is not None:
                    job.close()
                elif os.name != "nt":
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                elif process.poll() is None:
                    process.kill()
                process.wait(timeout=5)
        # A fast child can finish between polls: enforce limits before publishing.
        if sum(p.stat().st_size for p in root.rglob("*") if p.is_file()) > (
            self.config.max_upload_mb * 4 * 1048576
        ):
            reason = "output_limit"
        if self._storage_bytes() > self.config.max_storage_mb * 1048576:
            reason = "storage_limit"
        if reason or process.returncode or not (root / "result.json").exists():
            with (root / "stderr.txt").open("rb") as errors:
                diagnostic = errors.read(16384).decode("utf-8", "replace")
            from linexcel.execution import _failure_details

            failure = _failure_details("crashed", process.returncode, diagnostic)
            return {
                "error": {
                    "kind": reason or failure["kind"],
                    "message": reason or failure["summary"],
                    "diagnostic": ""
                    if task["operation"] in AI_OPERATIONS
                    else diagnostic,
                    "exitCode": process.returncode,
                }
            }
        return _read(root / "result.json")


class WebApp:
    def __init__(self, config: WebConfig, resolve_user: Callable | None = None):
        self.config = config
        self.store = TaskStore(config)
        self.resolve_user = resolve_user

    def close(self):
        self.store.close()

    def __call__(self, environ: dict, start_response: Callable):
        headers = [
            ("Cache-Control", "no-store"),
            ("X-Content-Type-Options", "nosniff"),
            ("Referrer-Policy", "no-referrer"),
        ]
        status = 200
        content_type = "application/json; charset=utf-8"
        try:
            path = environ.get("PATH_INFO", "/") or "/"
            mount = self.config.mount_path
            # WSGI hosts may already consume SCRIPT_NAME when mounting an app.
            if mount and environ.get("SCRIPT_NAME", "").rstrip("/") != mount:
                if path == mount:
                    path = "/"
                elif path.startswith(mount + "/"):
                    path = path[len(mount) :]
                else:
                    raise APIError(404, "Page introuvable")
            method = environ.get("REQUEST_METHOD", "GET")
            if self.resolve_user is not None:
                identity = self.resolve_user(environ)
                if not isinstance(identity, str) or not identity:
                    raise APIError(401, "Authentification requise")
            else:
                cookies = SimpleCookie()
                cookies.load(environ.get("HTTP_COOKIE", ""))
                cookie = cookies.get("linexcel_session")
                identity = cookie.value if cookie else ""
                if not re.fullmatch(r"[a-f0-9]{64}", identity):
                    identity = secrets.token_hex(32)
                    secure = (
                        "; Secure" if environ.get("wsgi.url_scheme") == "https" else ""
                    )
                    headers.append(
                        (
                            "Set-Cookie",
                            f"linexcel_session={identity}; "
                            f"Path={mount or '/'}; HttpOnly; SameSite=Strict; "
                            f"Max-Age=604800{secure}",
                        )
                    )
            owner = hashlib.sha256(identity.encode()).hexdigest()
            if method not in {"GET", "HEAD"}:
                origin = environ.get("HTTP_ORIGIN")
                if origin and urlsplit(origin).netloc != environ.get("HTTP_HOST"):
                    raise APIError(403, "Origine refusée")
                if environ.get("CONTENT_TYPE", "").split(";")[0] != "application/json":
                    raise APIError(415, "Content-Type application/json requis")
            if path == "/" and method == "GET":
                # Ensure relative API URLs remain beneath the configured mount.
                if not environ.get("PATH_INFO", "/").endswith("/"):
                    status = 303
                    headers.append(
                        ("Location", (mount or environ.get("SCRIPT_NAME", "")) + "/")
                    )
                    payload = b""
                else:
                    content_type = "text/html; charset=utf-8"
                    payload = (
                        Path(__file__).parent / "assets" / "app.html"
                    ).read_bytes()
                    headers.append(
                        (
                            "Content-Security-Policy",
                            "default-src 'self'; "
                            "script-src 'self' 'unsafe-inline'; "
                            "style-src 'self' 'unsafe-inline'; "
                            "img-src 'self' data: blob:; object-src 'none'; "
                            "frame-ancestors 'self'; base-uri 'none'",
                        )
                    )
            elif path == "/assets/cytoscape.min.js" and method in {"GET", "HEAD"}:
                # Fixed packaged asset only: never resolve a caller-supplied path.
                content_type = "text/javascript; charset=utf-8"
                payload = (
                    Path(__file__).parent / "assets" / "cytoscape.min.js"
                ).read_bytes()
            else:
                body = {}
                if method == "POST":
                    try:
                        length = int(environ.get("CONTENT_LENGTH") or 0)
                    except ValueError:
                        raise APIError(400, "Taille de requête invalide") from None
                    if not 0 <= length <= self.config.max_upload_mb * 1400000 + 65536:
                        raise APIError(413, "Requête trop volumineuse")
                    try:
                        body = json.loads(environ["wsgi.input"].read(length) or b"{}")
                    except (ValueError, UnicodeError):
                        raise APIError(400, "JSON invalide") from None
                    if not isinstance(body, dict):
                        raise APIError(400, "Objet JSON requis")
                payload = json.dumps(
                    self._route(owner, method, path, body),
                    ensure_ascii=False,
                    allow_nan=False,
                ).encode("utf-8")
        except APIError as exc:
            status = exc.status
            payload = json.dumps({"error": {"message": str(exc)}}).encode()
        except Exception:
            logging.getLogger(__name__).exception("Web request failed")
            status = 500
            payload = b'{"error":{"message":"Erreur interne du serveur"}}'
        from http import HTTPStatus

        headers.extend(
            [("Content-Type", content_type), ("Content-Length", str(len(payload)))]
        )
        start_response(f"{status} {HTTPStatus(status).phrase}", headers)
        return [b"" if environ.get("REQUEST_METHOD") == "HEAD" else payload]

    def ai_config(self) -> dict:
        from importlib.util import find_spec

        from linexcel.i18n import LANGUAGES

        installed = find_spec("openai") is not None
        return {
            "available": self.config.ai_enabled and installed,
            "configured": self.config.ai_enabled,
            "dependencyInstalled": installed,
            "model": self.config.ai_model if self.config.ai_enabled else None,
            "visionModel": (self.config.ai_vision_model or self.config.ai_model)
            if self.config.ai_enabled
            else None,
            "languages": list(LANGUAGES),
            "defaultSeconds": 600,
            "tokenBudget": self.config.ai_token_budget,
        }

    def _route(self, owner: str, method: str, path: str, body: dict) -> Any:
        parts = path.strip("/").split("/")
        if parts == ["api", "config"] and method == "GET":
            return {
                "ai": self.ai_config(),
                "limits": {
                    "memoryMb": self.config.total_memory_mb,
                    "uploadMb": self.config.max_upload_mb,
                    "retentionDays": self.config.retention_days,
                },
            }
        if parts == ["api", "projects"]:
            if method == "GET":
                return {"projects": self.store.list_projects(owner)}
            if method == "POST":
                return self.store.create(owner, body)
        if len(parts) >= 3 and parts[:2] == ["api", "projects"]:
            project_id = parts[2]
            if len(parts) == 3 and method == "GET":
                return self.store.snapshot(owner, project_id)
            if len(parts) == 3 and method == "DELETE":
                self.store.delete(owner, project_id)
                return {}
            if parts[3:] == ["export"] and method == "GET":
                return self.store.snapshot(owner, project_id)
            if parts[3:] == ["tasks"] and method == "POST":
                return {
                    "task": self.store.submit(
                        owner,
                        project_id,
                        body.get("operation", ""),
                        body.get("nodeId"),
                        body.get("budget", {}),
                        body.get("force", False),
                        body.get("options"),
                    )
                }
        if len(parts) >= 3 and parts[:2] == ["api", "tasks"]:
            if len(parts) == 3 and method == "GET":
                return {"task": self.store.public(self.store.get_task(owner, parts[2]))}
            if parts[3:] == ["cancel"] and method == "POST":
                return {"task": self.store.cancel(owner, parts[2])}
        raise APIError(404, "Ressource introuvable")


def create_app(
    data_dir: str | Path, *, resolve_user: Callable | None = None, **options: Any
) -> WebApp:
    """Create a WSGI app; host auth returns a stable, nonempty user ID or None."""
    return WebApp(WebConfig(Path(data_dir), **options), resolve_user)


def serve(args) -> int:
    from socketserver import ThreadingMixIn
    from wsgiref.simple_server import WSGIServer, make_server

    class Server(ThreadingMixIn, WSGIServer):
        daemon_threads = True

    app = create_app(
        args.data_dir,
        mount_path=args.mount_path,
        max_workers=args.workers,
        total_memory_mb=args.memory_mb,
        retention_days=args.retention_days,
        ai_enabled=args.ai,
        ai_base_url=args.ai_base_url,
        ai_model=args.ai_model,
        ai_vision_model=args.ai_vision_model,
        ai_token_budget=args.ai_token_budget,
        ai_api_key=os.getenv("LINEXCEL_AI_API_KEY"),
    )
    try:
        with make_server("127.0.0.1", args.port, app, server_class=Server) as server:
            print(
                f"Linexcel : http://127.0.0.1:{server.server_port}{args.mount_path}/",
                flush=True,
            )
            server.serve_forever()
    finally:
        app.close()
    return 0
