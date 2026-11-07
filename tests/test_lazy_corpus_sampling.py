"""The public harness must keep its sample stable and cache verdicts honest."""

import importlib.util
import math
from pathlib import Path

import pytest

from linexcel.values import readings_agree

MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts/validate_lazy_corpus.py"
spec = importlib.util.spec_from_file_location("validate_lazy_corpus", MODULE_PATH)
assert spec and spec.loader
corpus = importlib.util.module_from_spec(spec)
spec.loader.exec_module(corpus)


def test_quantiles_include_extremes_and_are_reproducible():
    assert corpus.spread(list(range(100)), 5) == [0, 25, 50, 74, 99]
    assert corpus.spread([1, 2], 5) == [1, 2]


def test_nodes_spread_across_sheets_and_signatures():
    nodes = [
        {"id": f"{sheet}!A{i}", "sheet": sheet, "kind": "cell", "formula": formula}
        for sheet in ("A", "B", "C")
        for i, formula in enumerate(("=1+1", "=SUM(B1:B3)", "=1+1"), 1)
    ]
    selected = corpus.sample_nodes(nodes, 6)
    assert [node["sheet"] for node in selected] == ["A", "B", "C", "A", "B", "C"]
    assert len({node["id"] for node in selected}) == 6
    assert corpus.sample_nodes(nodes[::-1], 6) == selected


def test_cache_comparison_respects_unknown_volatile_and_numeric_tolerance():
    node = {"cache_status": "saved", "cached_value": 2.0}
    result = {"status": "completed", "value": 2.0}
    assert corpus.comparison(node, result) == "exact_agreement"
    assert (
        corpus.comparison(node, {**result, "value": 2.0000000001})
        == "numeric_tolerance_agreement"
    )
    assert (
        corpus.comparison(node, {**result, "value": True})
        == "cache_difference_requires_review"
    )
    assert (
        corpus.comparison(node, {**result, "volatile": True})
        == "volatile_not_comparable"
    )
    assert (
        corpus.comparison({**node, "cache_status": "unknown"}, result)
        == "cache_unknown"
    )
    assert corpus.comparison(node, {"status": "unsupported"}) == "not_calculated"
    assert (
        corpus.comparison(node, {**result, "value": "display", "comparisonValue": 2.0})
        == "exact_agreement"
    )


def test_evaluation_exception_does_not_count_another_import():
    summary = {"imports": {"succeeded": 1}}
    report = {}
    corpus.record_exception(
        summary, report, "evaluation", RuntimeError("worker failed")
    )
    assert summary["imports"] == {"succeeded": 1}
    assert summary["harness_errors"] == {"evaluation": 1}
    assert report["exception"]["stage"] == "evaluation"


@pytest.mark.parametrize(
    ("saved", "value", "expected"),
    [
        (0.0, 5e-9, "cache_difference_requires_review"),
        (0.0, -5e-9, "cache_difference_requires_review"),
        (0.0, 1e-9, "numeric_tolerance_agreement"),
        (0.0, -1e-9, "numeric_tolerance_agreement"),
        (0.0, math.nextafter(1e-9, math.inf), "cache_difference_requires_review"),
        (0.3, 0.1 + 0.2, "numeric_tolerance_agreement"),
        (1e9, 1e9 + 0.5, "numeric_tolerance_agreement"),
        (1e9, 1e9 + 2.0, "cache_difference_requires_review"),
        (1, 1.0, "numeric_tolerance_agreement"),
        (0.0, 0.0, "exact_agreement"),
    ],
)
@pytest.mark.parametrize("reverse", [False, True])
def test_numeric_verdict_uses_product_tolerance(saved, value, expected, reverse):
    if reverse:
        saved, value = value, saved
    node = {"cache_status": "saved", "cached_value": saved}
    result = {"status": "completed", "value": value}
    assert corpus.comparison(node, result) == expected
    assert (expected != "cache_difference_requires_review") == (
        readings_agree(value, saved, None) == "same"
    )


def test_numeric_tolerance_uses_comparison_values_before_display_values():
    node = {
        "cache_status": "saved",
        "cached_value": "0.00",
        "cached_comparison_value": 0.0,
    }
    result = {"status": "completed", "value": "0.00", "comparisonValue": 5e-9}
    assert corpus.comparison(node, result) == "cache_difference_requires_review"
