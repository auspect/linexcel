"""Native stack safety must be tested across a process boundary."""

import subprocess
import sys
from types import SimpleNamespace

import pytest

from linexcel.engine import _evaluate_targets, _lexical_depth_bound


@pytest.mark.parametrize(
    "expression, rejected",
    [
        ('"=" + "+".join(["1"] * 3000)', True),
        ('"=" + "+".join(["1"] * 700)', False),
        ('"=SUM(" + ",".join(["1"] * 3000) + ")"', False),
        ('"=" + "-" * 3000 + "1"', True),
        ('"=1" + "%" * 3000', True),
        ('"=" + "^".join(["1"] * 3000)', True),
        ('"=" + "(" * 3000 + "1" + ")" * 3000', True),
        ("'=LEN(\"' + \"+-^(),\" * 3000 + '\")'", False),
        ('"=(" + ",".join(["A1"] * 3000) + ")"', True),
    ],
)
def test_guard_and_native_conversion_survive_in_subprocess(expression, rejected):
    script = f"""
from linexcel.engine import ast_depth, is_too_deep
formula = {expression}
assert is_too_deep(formula) is {rejected!r}
depth = ast_depth(formula)
assert (depth is None) is {rejected!r}, depth
print('survived')
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr
    assert "survived" in process.stdout


def test_quoted_content_and_structured_references_are_atomic():
    assert _lexical_depth_bound('=LEN("a""' + "+()," * 3000 + '")') < 10
    assert _lexical_depth_bound("='O''Brien " + "+()," * 3000 + "'!A1") < 10
    assert _lexical_depth_bound("=Table[[#Headers],[" + "+()," * 3000 + "]]") < 10


@pytest.mark.parametrize("targeted", [False, True])
@pytest.mark.parametrize("sheet_name", ["S", "R&D", "O'Brien"])
def test_full_analysis_preserves_cache_and_quarantines_before_import(
    targeted, sheet_name
):
    script = f"""
import io, zipfile
from openpyxl import Workbook
from linexcel import analyze
from linexcel.execution import ExecutionPolicy
sheet_name = {sheet_name!r}
w = Workbook(); w.active.title = sheet_name
w.active['A1'] = '=' + '+'.join(['1'] * 3000)
w.active['B1'] = '=2+3'
buffer = io.BytesIO(); w.save(buffer)
output = io.BytesIO()
with zipfile.ZipFile(buffer) as src, zipfile.ZipFile(output, 'w') as dst:
    for name in src.namelist():
        content = src.read(name)
        if name == 'xl/worksheets/sheet1.xml':
            content = content.replace(b'</f><v />', b'</f><v>3000</v>', 1)
        dst.writestr(name, content)
r = analyze(output.getvalue(), targets=["'"+sheet_name.replace("'","''")+"'!A1",
            "'"+sheet_name.replace("'","''")+"'!B1"] if {targeted!r} else None,
            execution=ExecutionPolicy(isolated=False))
nodes = {{n.get('addr'): n for n in r.nodes}}
assert nodes['A1']['value'] == 3000, nodes['A1']
assert nodes['A1']['valueSource'] == 'file'
assert nodes['B1']['value'] == 5
assert any('complexity guard' in warning for warning in r.warnings)
print('survived')
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr


def test_large_flat_sum_and_700_term_chain_still_recalculate():
    script = """
import io
from openpyxl import Workbook
from linexcel.engine import boot_engine
w = Workbook(); w.active.title = 'S'
w.active['A1'] = '=SUM(' + ','.join(['1'] * 3000) + ')'
w.active['B1'] = '=' + '+'.join(['1'] * 700)
output=io.BytesIO(); w.save(output)
session=boot_engine(output.getvalue(), [])
assert not session.quarantined
assert session.engine.get_value('S',1,1) == 3000
assert session.engine.get_value('S',1,2) == 700
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr


def test_target_limit_is_enforced_before_any_native_plan_or_evaluation():
    class Engine:
        def trace(self, *args, **kwargs):
            return SimpleNamespace(
                nodes={"S!A1": {}, "S!A999": {}},
                truncation=SimpleNamespace(incomplete=False),
            )

        def get_eval_plan(self, *args):
            pytest.fail("Rejected closure must not build a native evaluation plan")

        def evaluate_cells(self, *args):
            pytest.fail("Rejected closure must not be evaluated")

    with pytest.raises(ValueError, match="Target evaluation was not started"):
        _evaluate_targets(
            Engine(),
            {"S"},
            [("S", 999, 1)],
            [],
            SimpleNamespace(step=lambda _: None),
            max_cells_per_sheet=1,
        )
