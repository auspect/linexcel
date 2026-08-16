"""Build-time version, read by hatchling via ``[tool.hatch.version]``.

The git tag is the single source of truth. Sitting exactly on a clean ``v*``
tag is a release and yields that number; anything else is a development build
and carries a local segment saying how far past the tag it is::

    v1.2.3, clean       ->  1.2.3
    8 commits later     ->  1.2.3+dev.8.g7f5caf5
    with local edits    ->  1.2.3+dev.8.g7f5caf5.dirty

The local segment is not decoration: PyPI rejects local versions, so a build
that is not exactly a tag cannot reach the index even by accident.

Building from an sdist needs no git — hatchling records the number computed
here in ``PKG-INFO`` and reuses it. Only a source tree with no repository at
all, such as a "Download ZIP", falls back to ``0.0.0+unknown``.
"""

from __future__ import annotations

import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
UNKNOWN = "0.0.0+unknown"

# v1.2.3-8-g7f5caf5, with a -dirty suffix when the tree has uncommitted changes
_DESCRIBE = re.compile(
    r"^v?(?P<tag>[0-9]+\.[0-9]+\.[0-9]+)"
    r"-(?P<distance>[0-9]+)"
    r"-(?P<sha>g[0-9a-f]+)"
    r"(?P<dirty>-dirty)?$"
)


def _describe() -> str | None:
    try:
        completed = subprocess.run(
            ["git", "describe", "--tags", "--long", "--dirty", "--match", "v[0-9]*"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    return completed.stdout.strip() if completed.returncode == 0 else None


def compute_version() -> str:
    described = _describe()
    match = _DESCRIBE.match(described) if described else None
    if match is None:
        # No git, no tag, or a description this does not know how to read.
        # A release cannot land here unnoticed: the publish workflow compares
        # the built version against the tag before anything is uploaded.
        return UNKNOWN
    if match["distance"] == "0" and not match["dirty"]:
        return match["tag"]
    local = f"dev.{match['distance']}.{match['sha']}"
    if match["dirty"]:
        local += ".dirty"
    return f"{match['tag']}+{local}"
