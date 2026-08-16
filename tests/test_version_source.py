"""The build-time version derivation (version_source.py, at the repo root).

Only exercised at build time, so a regression here would otherwise surface as
a wrong number on a release rather than as a failing test.
"""

import importlib.util
from pathlib import Path

import pytest

_PATH = Path(__file__).resolve().parents[1] / "version_source.py"
_spec = importlib.util.spec_from_file_location("_version_source", _PATH)
version_source = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(version_source)


@pytest.fixture()
def described(monkeypatch):
    def _set(value):
        monkeypatch.setattr(version_source, "_describe", lambda: value)

    return _set


@pytest.mark.parametrize("tag", ["v1.2.3-0-g7f5caf5", "1.2.3-0-g7f5caf5"])
def test_sitting_on_a_clean_tag_is_a_release(described, tag):
    """No local segment, so the wheel is publishable and matches the tag."""
    described(tag)
    assert version_source.compute_version() == "1.2.3"


@pytest.mark.parametrize(
    ("describe", "expected"),
    [
        ("v1.2.3-8-g7f5caf5", "1.2.3+dev.8.g7f5caf5"),
        ("v1.2.3-1-g7f5caf5", "1.2.3+dev.1.g7f5caf5"),
        ("v1.2.3-8-g7f5caf5-dirty", "1.2.3+dev.8.g7f5caf5.dirty"),
    ],
)
def test_past_the_tag_is_a_dev_build(described, describe, expected):
    described(describe)
    assert version_source.compute_version() == expected


def test_a_dirty_tree_on_the_tag_is_not_a_release(described):
    """Uncommitted edits must not be publishable as the release itself."""
    described("v1.2.3-0-g7f5caf5-dirty")
    assert version_source.compute_version() == "1.2.3+dev.0.g7f5caf5.dirty"


@pytest.mark.parametrize("describe", [None, "", "not-a-version", "v1.2-3-gabc"])
def test_unusable_describe_degrades_rather_than_raising(described, describe):
    described(describe)
    assert version_source.compute_version() == "0.0.0+unknown"
