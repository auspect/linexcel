"""The targeted chain warning follows native formula dependencies."""

import formualizer as fz

from linexcel.engine import CHAIN_LAYERS_WARNING, _trace_formula_depth


def _trace(rows: list[list[str | float]], target: str):
    workbook = fz.Workbook()
    sheet = workbook.sheet("S")
    for row, values in enumerate(rows, 1):
        for column, value in enumerate(values, 1):
            if isinstance(value, str) and value.startswith("="):
                sheet.set_formula(row, column, value)
            else:
                sheet.set_value(row, column, value)
    return workbook.trace(
        [target],
        direction=fz.TraceDirection.Precedents,
        max_depth=100,
        max_nodes=1000,
        max_links=2000,
        max_work=5000,
        range_member_budget=100,
    )


def test_trace_depth_counts_a_long_formula_chain():
    rows = [[1]] + [[f"=A{row - 1}+1"] for row in range(2, 41)]
    trace = _trace(rows, "S!A40")

    # The formula source has 39 steps from A2 through A40. The trace walk
    # stops once it proves the warning threshold, so it returns that bound.
    assert _trace_formula_depth(trace) == CHAIN_LAYERS_WARNING


def test_trace_depth_proves_a_long_chain_before_exhausting_its_walk_budget():
    rows = [[1]] + [[f"=A{row - 1}+1"] for row in range(2, 41)]
    trace = _trace(rows, "S!A40")

    # A 25-step path proves the warning without visiting all 39 formulas.
    assert _trace_formula_depth(trace, work_limit=25) == CHAIN_LAYERS_WARNING


def test_trace_depth_distinguishes_a_wide_range_from_a_chain():
    rows = [["=1"] for _ in range(40)]
    rows.append(["=SUM(A1:A40)"])
    trace = _trace(rows, "S!A41")

    # A41 depends on 40 parallel one-step formulas, so its longest path has
    # two formula calculations even though the trace contains 41 formula cells.
    assert _trace_formula_depth(trace) == 2


def test_trace_depth_returns_unknown_when_its_walk_exceeds_budget():
    trace = _trace([[1], ["=A1+1"], ["=A2+1"], ["=A3+1"]], "S!A4")

    assert _trace_formula_depth(trace, work_limit=1) is None


def test_trace_depth_skips_cycle_edges():
    trace = _trace([["=B1+1", "=A1+1"]], "S!A1")

    # The trace contains both formulas, but its cycle disposition does not
    # expand the return edge; the path remains finite.
    assert _trace_formula_depth(trace) == 2
