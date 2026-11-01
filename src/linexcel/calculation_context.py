"""Replayable clock/random context and source-declared iteration settings."""

import math
import secrets
from datetime import UTC, datetime


def iteration_settings(element):
    settings = {
        "enabled": False,
        "maxIterations": 100,
        "maxChange": 0.001,
        "valid": True,
    }
    if element is None:
        return settings
    try:
        enabled = element.get("iterate", "0").lower()
        count = int(element.get("iterateCount", "100"))
        change = float(element.get("iterateDelta", "0.001"))
        if (
            enabled not in {"0", "1", "false", "true"}
            or count < 1
            or not math.isfinite(change)
            or change < 0
        ):
            raise ValueError("Invalid workbook iteration settings")
        settings.update(
            enabled=enabled in {"1", "true"}, maxIterations=count, maxChange=change
        )
    except (ValueError, TypeError):
        settings["valid"] = False
    return settings


def evaluation_context(context=None):
    if context is None:
        return {
            "timestampUtc": datetime.now(UTC).isoformat(),
            "timezone": "utc",
            "seed": secrets.randbits(64),
        }
    if not isinstance(context, dict) or set(context) != {
        "timestampUtc",
        "timezone",
        "seed",
    }:
        raise ValueError("Evaluation context requires timestampUtc, timezone and seed")
    seed = context["seed"]
    if isinstance(seed, bool) or not isinstance(seed, int) or not 0 <= seed < 2**64:
        raise ValueError("Evaluation seed must be an unsigned 64-bit integer")
    if context["timezone"] != "utc":
        raise ValueError("Replay context currently requires the explicit UTC timezone")
    instant = datetime.fromisoformat(context["timestampUtc"])
    offset = instant.utcoffset()
    if offset is None or offset.total_seconds() != 0:
        raise ValueError("Evaluation timestamp must include a UTC offset")
    return {**context, "timestampUtc": instant.isoformat()}
