"""Explicit dependency diagnostics; never send workbook data or generate text."""

from __future__ import annotations

import importlib.util
import shutil
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from types import SimpleNamespace
from typing import Any


def directory_health(config) -> dict:
    """Read-only CLI probe, including while a server owns the workspace lock."""
    root = Path(config.data_dir).resolve()

    def size():
        total = 0
        for path in root.rglob("*"):
            try:
                if path.is_file():
                    total += path.stat().st_size
            except OSError:
                pass
        return total

    return health(SimpleNamespace(root=root, config=config, _storage_bytes=size))


def health(store) -> dict:
    from linexcel.insights import find_libreoffice, find_pdftoppm

    checks: dict[str, dict[str, Any]] = {}
    for name in ("formualizer", "openpyxl", "openai", "playwright", "pillow"):
        try:
            checks[name] = {"status": "ready", "version": version(name)}
        except PackageNotFoundError:
            checks[name] = {"status": "unavailable", "code": "package_missing"}
    for name, executable in (
        ("libreoffice", find_libreoffice()),
        ("poppler", find_pdftoppm()),
    ):
        checks[name] = {
            "status": "ready" if executable else "unavailable",
            "code": "executable_found" if executable else "executable_missing",
        }
    used = store._storage_bytes()
    quota = store.config.max_storage_mb * 1048576
    location = store.root
    while not location.exists() and location != location.parent:
        location = location.parent
    free = shutil.disk_usage(location).free
    checks["storage"] = {
        "status": "ready" if min(quota - used, free) > 0 else "unavailable",
        "usedBytes": used,
        "quotaBytes": quota,
        "freeBytes": free,
    }
    config = store.config
    checks["ai"] = {"status": "disabled"}
    if config.ai_enabled:
        checks["ai"] = {
            "status": "unavailable",
            "code": "client_missing",
            "model": config.ai_model,
            "visionModel": config.ai_vision_model or config.ai_model,
        }
        if (
            importlib.util.find_spec("httpx") is not None
            and checks["openai"]["status"] == "ready"
        ):
            import httpx

            headers = {}
            if config.ai_api_key:
                headers["Authorization"] = "Bearer " + config.ai_api_key
            try:
                with httpx.Client(timeout=3, follow_redirects=False) as client:
                    with client.stream(
                        "GET",
                        config.ai_base_url.rstrip("/") + "/models",
                        headers=headers,
                    ) as response:
                        response.raise_for_status()
                        raw = bytearray()
                        for chunk in response.iter_bytes():
                            raw.extend(chunk)
                            if len(raw) > 1048576:
                                raise ValueError(
                                    "Model listing exceeds diagnostic limit"
                                )
                import json

                payload = json.loads(raw)
                if not isinstance(payload, dict) or not isinstance(
                    payload.get("data"), list
                ):
                    raise ValueError("Invalid model listing")
                models = {item["id"] for item in payload["data"]}
                wanted = {config.ai_model, config.ai_vision_model or config.ai_model}
                missing = sorted(
                    m for m in wanted if m not in models and m + ":latest" not in models
                )
                checks["ai"].update(
                    status="ready" if not missing else "unavailable",
                    code="models_listed" if not missing else "model_missing",
                    missingModels=missing,
                    visionCapability="not_probed",
                )
            except (httpx.HTTPError, ValueError, KeyError, TypeError):
                checks["ai"]["code"] = "service_unreachable_or_listing_unavailable"
    return {"checks": checks, "generationRequested": False}
