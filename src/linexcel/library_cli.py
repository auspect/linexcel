"""Offline library commands use the same owner, data directory and writer lock."""

from __future__ import annotations

import json
import re
import sys

from linexcel.catalog import CatalogLimits
from linexcel.web import APIError, TaskStore, WebConfig
from linexcel.web_local import LocalProfile


def run(args) -> int:
    limits = CatalogLimits(
        entries_per_workbook=args.entries_per_workbook,
        catalog_bytes=args.catalog_mb * 1048576,
        build_seconds=args.build_seconds,
        connections=args.connections,
        hits=args.hits,
        response_bytes=args.response_kb * 1024,
    )
    store = TaskStore(
        WebConfig(args.data_dir, retention_days=None, catalog_limits=limits)
    )
    try:
        owner = args.owner or LocalProfile(store.root).owner
        if not re.fullmatch(r"[a-f0-9]{64}", owner):
            raise ValueError("owner must be a SHA-256 owner id")
        operation = args.library_action
        if operation == "rebuild":
            store.library.catalog(owner).ensure(
                store.library.sources(owner), force=True
            )
            operation = "status"
        names = (
            "query",
            "kinds",
            "project",
            "revision",
            "generation",
            "offset",
            "limit",
        )
        if operation == "lineage":
            names = ("handle", "depth", "limit")
        elif operation == "occurrences":
            names = ("group", "generation", "after", "limit")
        params = {
            name: [str(getattr(args, name))]
            for name in names
            if getattr(args, name, None) is not None
        }
        result = store.library.query(owner, operation, params, synchronous=True)
        payload = json.dumps(
            result, ensure_ascii=False, separators=(",", ":"), allow_nan=False
        )
        binary = getattr(sys.stdout, "buffer", None)
        if binary is None:  # In-process text capture/embedding, with no byte encoding.
            sys.stdout.write(payload)
        else:
            binary.write(payload.encode("utf-8"))
        if result.get("buildState", {}).get("status") == "failed":
            return 2
        coverage = result.get("coverage", {})
        return 3 if not coverage.get("complete", True) else 0
    except APIError as error:
        raise ValueError(str(error)) from None
    finally:
        store.close()
