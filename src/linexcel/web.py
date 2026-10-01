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
from urllib.parse import parse_qs, urlsplit

TERMINAL = {"succeeded", "failed", "cancelled"}
OPERATIONS = {"import", "evaluate", "capture", "document", "describe_capture"}
AI_OPERATIONS = {"document", "describe_capture"}
MAX_RESULT_BYTES = 16 * 1024 * 1024
MAX_PREVIEW_STEPS = 200


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
    max_upload_mb: int = 256
    max_storage_mb: int = 2048
    retention_days: int = 7
    max_pending_tasks: int = 32
    ai_enabled: bool = False
    ai_base_url: str = field(default="http://localhost:11434/v1", repr=False)
    ai_api_key: str | None = field(default=None, repr=False)
    ai_model: str = "qwen3.8"
    ai_vision_model: str | None = None
    ai_token_budget: int = 16000
    allow_local_import: bool = False

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
        if not isinstance(self.ai_enabled, bool) or not isinstance(
            self.allow_local_import, bool
        ):
            raise ValueError("ai_enabled and allow_local_import must be boolean")
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
        self._project_exports: dict[tuple[str, str], int] = {}
        self._revisions: dict[str, tuple[tuple, str]] = {}
        from importlib.metadata import version

        cache = hashlib.sha256(version("formualizer").encode())
        for name in (
            "lazy.py",
            "lazy_store.py",
            "engine.py",
            "excel_compat.py",
            "values.py",
            "loader.py",
            "external.py",
            "refs.py",
            "insights.py",
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
        from linexcel.web_upload import UploadStore

        self.uploads = UploadStore(self)

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
                if task.get("result") is not None:
                    self._save_result(task_path.parent, task, task["result"])
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
    def _filename(name: Any) -> str:
        if (
            not isinstance(name, str)
            or len(name) > 180
            or not re.fullmatch(r"[^/\\:\x00-\x1f]+\.(?:xlsx|xlsm)", name, re.I)
            or name.startswith(".")
            or name.rstrip(". ") != name
            or re.match(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])\.", name, re.I)
        ):
            raise APIError(400, "Nom de fichier .xlsx/.xlsm invalide")
        return name

    @staticmethod
    def _file(value: Any) -> tuple[str, bytes]:
        if not isinstance(value, dict):
            raise APIError(400, "Fichier invalide")
        name = TaskStore._filename(value.get("name", ""))
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
                used + self.uploads.remaining_bytes() + sum(len(d) for _, d in files)
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

    def _revision(self, path: Path) -> str:
        files = [path / "workbook.xlsx", *sorted((path / "references").iterdir())]
        signature = tuple(
            (file.name, info.st_size, info.st_mtime_ns, info.st_ctime_ns, info.st_ino)
            for file in files
            for info in (file.stat(),)
        )
        cached = self._revisions.get(str(path))
        if cached is not None and cached[0] == signature:
            return cached[1]
        digest = hashlib.sha256()
        for file in files:
            digest.update(file.name.encode("utf-8"))
            with file.open("rb") as stream:
                while chunk := stream.read(1024 * 1024):
                    digest.update(chunk)
        revision = digest.hexdigest()
        self._revisions[str(path)] = (signature, revision)
        return revision

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
                "graph": self._graph(path),
                "tasks": [
                    self.public(t, include_result=False)
                    for t in self.tasks.values()
                    if t["owner"] == owner and t["projectId"] == project_id
                ],
            }

    @staticmethod
    def _graph(path: Path) -> dict | None:
        if (path / "index.sqlite").is_file():
            from linexcel.lazy_store import graph_page

            return graph_page(path / "index.sqlite")
        return _read(path / "graph.json") if (path / "graph.json").is_file() else None

    @staticmethod
    def _node(path: Path, node_id: str) -> dict | None:
        if (path / "index.sqlite").is_file():
            from linexcel.lazy_store import get_node

            return get_node(path / "index.sqlite", node_id)
        graph = TaskStore._graph(path)
        return next(
            (n for n in (graph or {}).get("nodes", []) if n["id"] == node_id), None
        )

    def graph_query(self, owner: str, project_id: str, kind: str, query: dict) -> dict:
        """Read only bounded SQL results in the HTTP coordinator."""
        from linexcel.lazy_store import graph_neighborhood, graph_page, patterns_page

        def value(name: str, default: str = "") -> str:
            values = query.get(name, [default])
            if len(values) != 1 or len(values[0]) > 1024:
                raise APIError(400, "Paramètre de recherche invalide")
            return values[0]

        try:
            offset = int(value("offset", "0"))
            limit = int(value("limit", "200"))
            depth = int(value("depth", "1"))
        except ValueError:
            raise APIError(400, "Pagination invalide") from None
        if offset < 0 or not 1 <= limit <= 200 or depth not in {1, 2}:
            raise APIError(400, "Pagination : 1–200 éléments, profondeur 1 ou 2")
        with self.lock:
            path = self.project_path(owner, project_id)
            index = path / "index.sqlite"
            if not index.is_file():
                raise APIError(
                    409, "Actualisez la structure pour activer l’index sur disque"
                )
            if kind == "node":
                node = self._node(path, value("nodeId"))
                if node is None:
                    raise APIError(404, "Nœud introuvable")
                return {"node": node}
            if kind == "neighborhood":
                node_id = value("nodeId")
                if self._node(path, node_id) is None:
                    raise APIError(404, "Nœud introuvable")
                return graph_neighborhood(index, node_id, limit=limit, depth=depth)
            kwargs = {"offset": offset, "limit": limit, "sheet": value("sheet") or None}
            if kind == "patterns":
                return patterns_page(index, **kwargs)
            return graph_page(index, query=value("q") or None, **kwargs)

    @staticmethod
    def _save_result(root: Path, task: dict, result: dict) -> None:
        _write(root / "result-data.json", result)
        preview = result.copy()
        if (
            isinstance(result.get("steps"), list)
            and len(result["steps"]) > MAX_PREVIEW_STEPS
        ):
            preview["steps"] = result["steps"][:MAX_PREVIEW_STEPS]
            preview["stepsPagination"] = {
                "offset": 0,
                "limit": MAX_PREVIEW_STEPS,
                "total": len(result["steps"]),
                "hasMore": True,
            }
        _write(root / "result-preview.json", preview)
        task.update(result=None, resultAvailable=True)

    def public(self, task: dict, *, include_result: bool = True) -> dict:
        public = {k: v for k, v in task.items() if k != "owner"}
        if not include_result:
            public["result"] = None
        elif public.get("resultAvailable") and public.get("result") is None:
            public["result"] = self._task_result(task)
        return public

    def _task_result(self, task: dict) -> dict | None:
        if task.get("result") is not None:
            return task["result"]
        if not task.get("resultAvailable"):
            return None
        path = self.project_path(task["owner"], task["projectId"])
        return _read(path / "tasks" / task["id"] / "result-preview.json")

    def task_export(self, owner: str, task_id: str) -> Path:
        task = self.get_task(owner, task_id, include_result=False)
        path = self.project_path(owner, task["projectId"])
        result = path / "tasks" / task_id / "result-data.json"
        if not task.get("resultAvailable") or not result.is_file():
            raise APIError(404, "Résultat indisponible")
        return result

    def get_task(
        self, owner: str, task_id: str, *, include_result: bool = True
    ) -> dict:
        with self.lock:
            task = self.tasks.get(task_id)
            if task is None or task["owner"] != owner:
                raise APIError(404, "Opération introuvable")
            result = task.copy()
            if include_result:
                result["result"] = self._task_result(task)
            return result

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
            if operation == "import" and self._project_exports.get((owner, project_id)):
                raise APIError(409, "Un export est en cours ; réessayez après sa fin")
            if self._revision(path) != project["revision"]:
                raise APIError(
                    409, "Sources modifiées sur disque ; importez une nouvelle version"
                )
            pattern_id = None
            if operation == "evaluate" or (
                operation == "document" and node_id is not None
            ):
                if not isinstance(node_id, str) or not any(
                    (path / name).is_file() for name in ("index.sqlite", "graph.json")
                ):
                    raise APIError(409, "Importez le graphe avant le recalcul")
                node = self._node(path, node_id)
                if node is None:
                    raise APIError(404, "Nœud introuvable")
                pattern_id = node.get("patternId")
            ai_snapshot = None
            ai_cache_key = None
            if operation in AI_OPERATIONS:
                from linexcel.web_ai import capture_evidence, digest, evidence

                if operation == "document":
                    if not any(
                        (path / name).is_file()
                        for name in ("index.sqlite", "graph.json")
                    ):
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
                        and (t.get("resultAvailable") or t.get("result"))
                    ]
                    completed = [
                        self.get_task(owner, t["id"])
                        for t in sorted(
                            completed, key=lambda t: t["createdAt"], reverse=True
                        )[: 1 if node_id is not None else 4]
                    ]
                    if (path / "index.sqlite").is_file():
                        from linexcel.lazy_store import evidence_graph

                        graph = evidence_graph(path / "index.sqlite", node_id)
                    else:
                        graph = _read(path / "graph.json")
                    ai_snapshot = evidence(graph, node_id, completed)
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
                    if "data" not in shots[index]:
                        from linexcel.web_ai import MAX_IMAGE_BYTES

                        image_path = self.capture_path(
                            owner, project_id, capture_id, index
                        )
                        with image_path.open("rb") as image:
                            image_data = image.read(MAX_IMAGE_BYTES + 1)
                        if len(image_data) > MAX_IMAGE_BYTES:
                            raise APIError(413, "Capture supérieure à 8 Mio pour l’IA")
                        selected = {
                            **shots[index],
                            "data": "data:image/png;base64,"
                            + base64.b64encode(image_data).decode(),
                        }
                        # Only the explicitly selected image enters the AI dossier.
                        captured = {
                            **captured,
                            "result": {
                                **captured["result"],
                                "screenshots": [
                                    selected if i == index else shot
                                    for i, shot in enumerate(shots)
                                ],
                            },
                        }
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

    def capture_path(
        self,
        owner: str,
        project_id: str,
        task_id: str,
        index: int,
        thumbnail: bool = False,
    ) -> Path:
        with self.lock:
            project = self.project_path(owner, project_id)
            task = self.get_task(owner, task_id)
            if (
                task["projectId"] != project_id
                or task["operation"] != "capture"
                or task["status"] != "succeeded"
                or not 0
                <= index
                < len((task.get("result") or {}).get("screenshots", []))
            ):
                raise APIError(404, "Capture introuvable")
            filename = f"{index}.thumb.png" if thumbnail else f"{index}.png"
            path = project / "captures" / task_id / filename
            path.resolve().relative_to(project.resolve())
            if not path.is_file():
                raise APIError(404, "Capture introuvable")
            return path

    def delete(self, owner: str, project_id: str):
        with self.lock:
            path = self.project_path(owner, project_id)
            if self._project_exports.get((owner, project_id)):
                raise APIError(409, "Un export est en cours ; réessayez après sa fin")
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
            self._revisions.pop(str(path), None)
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
                    payload: dict[str, Any] = result["result"]
                    if (
                        task["operation"] == "capture"
                        and (root / "capture-files").is_dir()
                    ):
                        project = root.parent.parent
                        screenshots = []
                        for index, shot in enumerate(payload.get("screenshots", [])):
                            if shot.get("file") != f"{index}.png":
                                raise ValueError("Invalid capture filename")
                            image_path = root / "capture-files" / f"{index}.png"
                            with image_path.open("rb") as source:
                                if source.read(8) != b"\x89PNG\r\n\x1a\n":
                                    raise ValueError("Invalid PNG capture")
                            metadata = {
                                "sheet": shot.get("sheet", ""),
                                "name": shot.get("name", f"{index}.png"),
                                "url": f"api/projects/{task['projectId']}/captures/"
                                f"{task['id']}/{index}",
                            }
                            if shot.get("thumbnail") == f"{index}.thumb.png":
                                thumb = root / "capture-files" / shot["thumbnail"]
                                if thumb.is_file():
                                    metadata["thumbnailUrl"] = (
                                        metadata["url"] + "?thumbnail=1"
                                    )
                            screenshots.append(metadata)
                        capture_root = project / "captures"
                        capture_root.mkdir(exist_ok=True)
                        (root / "capture-files").rename(capture_root / task["id"])
                        payload = {**payload, "screenshots": screenshots}
                    if task["operation"] == "import":
                        project = root.parent.parent
                        if (root / "index.sqlite").is_file():
                            (root / "index.sqlite").replace(project / "index.sqlite")
                            (project / "graph.json").unlink(missing_ok=True)
                        else:
                            _write(project / "graph.json", payload)
                        payload = {
                            "nodeCount": payload.get("meta", {}).get(
                                "nodeCount", len(payload.get("nodes", []))
                            ),
                            "indexStatus": payload.get("meta", {}).get(
                                "status", "complete"
                            ),
                            "diagnostics": payload.get("meta", {}).get(
                                "diagnostics", []
                            ),
                        }
                    self._save_result(root, task, payload)
                    task.update(status="succeeded")
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
                    for scratch in (
                        "result.json",
                        "request.json",
                        "ai_input.json",
                        "index.sqlite",
                        "index.sqlite-journal",
                        "index.sqlite-wal",
                        "index.sqlite-shm",
                    ):
                        (root / scratch).unlink(missing_ok=True)
                    for scratch in root.glob("index.sqlite.*.building*"):
                        scratch.unlink(missing_ok=True)
                    if (root / "captures").exists():
                        shutil.rmtree(root / "captures")
                    if (root / "capture-files").exists():
                        shutil.rmtree(root / "capture-files")
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
        # Disk-backed indexes may be much larger than the compressed source.
        # Enforce the explicit storage quota, never a multiple of upload size.
        if self._storage_bytes() > self.config.max_storage_mb * 1048576:
            reason = "storage_limit"
        if reason or process.returncode or not (root / "result.json").exists():
            with (root / "stderr.txt").open("rb") as errors:
                diagnostic = errors.read(16384).decode("utf-8", "replace")
            from linexcel.execution import _failure_details

            failure = _failure_details("crashed", process.returncode, diagnostic)
            kind = reason or failure["kind"]
            message = reason or failure["summary"]
            if kind == "memory_limit":
                message = (
                    f"Budget mémoire de {request['memoryMb']} Mio insuffisant pour "
                    "cette opération. Augmentez le budget puis réessayez ; "
                    "le classeur et les autres résultats restent accessibles."
                )
            return {
                "error": {
                    "kind": kind,
                    "message": message,
                    "diagnostic": ""
                    if task["operation"] in AI_OPERATIONS
                    else diagnostic,
                    "exitCode": process.returncode,
                }
            }
        if (root / "result.json").stat().st_size > MAX_RESULT_BYTES:
            return {
                "error": {
                    "kind": "result_limit",
                    "message": "Résultat trop volumineux ; sélectionnez une cellule "
                    "avec moins de dépendances. Les autres opérations "
                    "restent disponibles.",
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
            upload_match = re.fullmatch(
                r"/api/uploads/([a-f0-9]{32})/files/([0-9]+)", path
            )
            raw_upload = method == "PUT" and upload_match is not None
            if method not in {"GET", "HEAD"}:
                origin = environ.get("HTTP_ORIGIN")
                if origin and (
                    urlsplit(origin).netloc != environ.get("HTTP_HOST")
                    or urlsplit(origin).scheme != environ.get("wsgi.url_scheme")
                ):
                    raise APIError(403, "Origine refusée")
                expected_type = (
                    "application/octet-stream" if raw_upload else "application/json"
                )
                if environ.get("CONTENT_TYPE", "").split(";")[0] != expected_type:
                    raise APIError(415, f"Content-Type {expected_type} requis")
            if path == "/api/projects/local" and method == "POST":
                import ipaddress

                try:
                    local_peer = ipaddress.ip_address(
                        environ.get("REMOTE_ADDR", "")
                    ).is_loopback
                except ValueError:
                    local_peer = False
                host = urlsplit("//" + environ.get("HTTP_HOST", "")).hostname
                if (
                    not self.config.allow_local_import
                    or self.resolve_user is not None
                    or not local_peer
                    or host not in {"localhost", "127.0.0.1", "::1"}
                ):
                    raise APIError(403, "Import local réservé au serveur loopback")
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
            elif (
                project_export := re.fullmatch(
                    r"/api/projects/([a-f0-9]{32})/export", path
                )
            ) and method in {"GET", "HEAD"}:
                from linexcel.web_export import ProjectExport

                export = ProjectExport(self.store, owner, project_export[1])
                try:
                    headers.extend(
                        [
                            ("Content-Type", "application/json; charset=utf-8"),
                            (
                                "Content-Disposition",
                                f'attachment; filename="{project_export[1]}.json"',
                            ),
                        ]
                    )
                    start_response("200 OK", headers)
                except BaseException:
                    export.close()
                    raise
                if method == "HEAD":
                    export.close()
                    return []
                return export
            elif (
                export_match := re.fullmatch(r"/api/tasks/([a-f0-9]{32})/export", path)
            ) and method in {"GET", "HEAD"}:
                from wsgiref.util import FileWrapper

                result_path = self.store.task_export(owner, export_match[1])
                stream = result_path.open("rb")
                headers.extend(
                    [
                        ("Content-Type", "application/json; charset=utf-8"),
                        ("Content-Length", str(os.fstat(stream.fileno()).st_size)),
                        (
                            "Content-Disposition",
                            f'attachment; filename="{export_match[1]}.json"',
                        ),
                    ]
                )
                start_response("200 OK", headers)
                if method == "HEAD":
                    stream.close()
                    return [b""]
                return FileWrapper(stream, 65536)
            elif method in {"GET", "HEAD"} and (
                capture_match := re.fullmatch(
                    r"/api/projects/([a-f0-9]{32})/captures/([a-f0-9]{32})/([0-9]+)",
                    path,
                )
            ):
                from wsgiref.util import FileWrapper

                image_path = self.store.capture_path(
                    owner,
                    capture_match[1],
                    capture_match[2],
                    int(capture_match[3]),
                    thumbnail=parse_qs(
                        environ.get("QUERY_STRING", ""), max_num_fields=12
                    ).get("thumbnail")
                    == ["1"],
                )
                try:
                    image_stream = image_path.open("rb")
                except FileNotFoundError:
                    raise APIError(404, "Capture introuvable") from None
                headers.extend(
                    [
                        ("Content-Type", "image/png"),
                        (
                            "Content-Length",
                            str(os.fstat(image_stream.fileno()).st_size),
                        ),
                    ]
                )
                start_response("200 OK", headers)
                if method == "HEAD":
                    image_stream.close()
                    return []
                return FileWrapper(image_stream, 65536)
            elif raw_upload:
                try:
                    length = int(environ.get("CONTENT_LENGTH") or "")
                    offset = int(environ.get("HTTP_X_UPLOAD_OFFSET") or "")
                except ValueError:
                    raise APIError(
                        400, "Taille et position du fragment requises"
                    ) from None
                payload = json.dumps(
                    self.store.uploads.append(
                        owner,
                        upload_match[1],
                        int(upload_match[2]),
                        offset,
                        length,
                        environ["wsgi.input"],
                    ),
                    ensure_ascii=False,
                ).encode("utf-8")
            else:
                body = {}
                if method == "POST":
                    try:
                        length = int(environ.get("CONTENT_LENGTH") or 0)
                    except ValueError:
                        raise APIError(400, "Taille de requête invalide") from None
                    # The compatibility JSON API is deliberately bounded; the UI
                    # streams workbook bytes through owner-scoped upload sessions.
                    limit = 8 * 1048576 if path == "/api/projects" else 65536
                    if not 0 <= length <= limit:
                        raise APIError(
                            413,
                            "JSON trop volumineux ; "
                            "utilisez le transfert par fragments",
                        )
                    try:
                        body = json.loads(environ["wsgi.input"].read(length) or b"{}")
                    except (ValueError, UnicodeError):
                        raise APIError(400, "JSON invalide") from None
                    if not isinstance(body, dict):
                        raise APIError(400, "Objet JSON requis")
                try:
                    query = parse_qs(environ.get("QUERY_STRING", ""), max_num_fields=12)
                except ValueError:
                    raise APIError(400, "Trop de paramètres de recherche") from None
                payload = json.dumps(
                    self._route(
                        owner,
                        method,
                        path,
                        body,
                        query,
                    ),
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

    def _route(
        self,
        owner: str,
        method: str,
        path: str,
        body: dict,
        query: dict | None = None,
    ) -> Any:
        parts = path.strip("/").split("/")
        if parts == ["api", "config"] and method == "GET":
            from linexcel.web_upload import CHUNK_BYTES

            return {
                "ai": self.ai_config(),
                "uploads": {
                    "chunkBytes": CHUNK_BYTES,
                    "localImport": self.config.allow_local_import
                    and self.resolve_user is None,
                },
                "limits": {
                    "memoryMb": self.config.total_memory_mb,
                    "uploadMb": self.config.max_upload_mb,
                    "retentionDays": self.config.retention_days,
                },
            }
        if parts == ["api", "uploads"] and method == "POST":
            return self.store.uploads.begin(owner, body)
        if len(parts) >= 3 and parts[:2] == ["api", "uploads"]:
            if len(parts) == 3 and method == "GET":
                return self.store.uploads.get(owner, parts[2])
            if len(parts) == 3 and method == "DELETE":
                return self.store.uploads.cancel(owner, parts[2])
            if parts[3:] == ["complete"] and method == "POST":
                return self.store.uploads.complete(owner, parts[2])
        if parts == ["api", "projects", "local"] and method == "POST":
            return self.store.uploads.local(owner, body)
        if parts == ["api", "projects"]:
            if method == "GET":
                return {"projects": self.store.list_projects(owner)}
            if method == "POST":
                return self.store.create(owner, body)
        if len(parts) >= 3 and parts[:2] == ["api", "projects"]:
            project_id = parts[2]
            if (
                len(parts) == 4
                and parts[3] in {"nodes", "node", "neighborhood", "patterns"}
                and method == "GET"
            ):
                return self.store.graph_query(owner, project_id, parts[3], query or {})
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
    from wsgiref.simple_server import WSGIRequestHandler, WSGIServer, make_server

    class Server(ThreadingMixIn, WSGIServer):
        daemon_threads = True

    class Handler(WSGIRequestHandler):
        def setup(self):
            self.request.settimeout(60)
            super().setup()

    app = create_app(
        args.data_dir,
        mount_path=args.mount_path,
        max_workers=args.workers,
        total_memory_mb=args.memory_mb,
        retention_days=args.retention_days,
        max_upload_mb=args.max_upload_mb,
        max_storage_mb=args.max_storage_mb,
        allow_local_import=not args.no_local_import,
        ai_enabled=args.ai,
        ai_base_url=args.ai_base_url,
        ai_model=args.ai_model,
        ai_vision_model=args.ai_vision_model,
        ai_token_budget=args.ai_token_budget,
        ai_api_key=os.getenv("LINEXCEL_AI_API_KEY"),
    )
    try:
        with make_server(
            "127.0.0.1", args.port, app, server_class=Server, handler_class=Handler
        ) as server:
            print(
                f"Linexcel : http://127.0.0.1:{server.server_port}{args.mount_path}/",
                flush=True,
            )
            server.serve_forever()
    finally:
        app.close()
    return 0
