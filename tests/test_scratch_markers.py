"""Private scratch markers cannot consume legitimate workbook strings."""

import pytest

from linexcel.decompose import SCRATCH_SENTINEL, _scratch_eval, _scratch_marker
from linexcel.loader import CachedValues
from linexcel.resolver import _Budget, _ValueResolver


class RefusingEngine:
    def __init__(self):
        self.formulas = {}

    def set_formula(self, sheet, row, col, formula):
        if "Refused" not in formula:
            self.formulas[sheet, row, col] = formula

    def evaluate_cell(self, sheet, row, col):
        formula = self.formulas[sheet, row, col].lstrip("=")
        return formula[1:-1] if formula.startswith('"') else 3

    def evaluate_cells(self, targets):
        return [self.evaluate_cell(*target) for target in targets]


def resolver(engine, **kwargs):
    return _ValueResolver(
        engine, {"S"}, CachedValues({}, set(), False), [], _Budget(100), True, **kwargs
    )


def test_each_scratch_operation_uses_a_distinct_marker():
    assert len({_scratch_marker() for _ in range(100)}) == 100


@pytest.mark.parametrize("batched", [False, True])
def test_formulas_can_return_the_old_marker_and_refusal_still_stays_unknown(batched):
    engine = RefusingEngine()
    expression = f'"{SCRATCH_SENTINEL}"'
    if batched:
        values = resolver(engine)
        values.preload_steps([expression], "S")
        assert values.eval_expr(expression, "S") == (SCRATCH_SENTINEL, True)
        values.preload_steps(["Refused()"], "S")
        assert values.eval_expr("Refused()", "S") == (None, False)
    else:
        assert _scratch_eval(engine, expression, "S") == (SCRATCH_SENTINEL, True)
        assert _scratch_eval(engine, "Refused()", "S") == (None, False)


@pytest.mark.parametrize(
    "expression", ["=B1+5", "=IFERROR(B1,5)", "=IF(ISERROR(B1),5,B1)"]
)
def test_recovery_limit_never_evaluates_a_formula_with_unresolved_precedents(
    expression,
):
    class BrokenEngine:
        def get_value(self, sheet, row, col):
            return None

        def get_formula(self, sheet, row, col):
            return "=1+2" if col == 2 else None

    values = resolver(BrokenEngine(), engine_alive=False, max_chain_depth=1)

    def scratch_must_not_run(*args):
        pytest.fail("A truncated recovery must not run dependent scratch formulas")

    values._eval_raw = scratch_must_not_run
    assert values._eval_formula("S", expression, 0) == (None, None)
    assert any("recovery stopped" in warning for warning in values.warnings)
    assert values._recover("S", 1, 3, expression) == (None, None)
    # GraphBuilder uses this flag to withhold decomposition as well as the
    # cell value; no scratch proof may reintroduce the incomplete result.
    assert ("S", 1, 3) in values.unavailable
