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
import formualizer as fz
from linexcel.engine import ast_depth, is_too_deep
formula = {expression}
parser_limit = False
if not {rejected!r}:
    try:
        fz.parse(formula)
    except Exception as error:
        parser_limit = "AST height limit exceeded" in str(error)
expected = {rejected!r} or parser_limit
assert is_too_deep(formula) is expected
depth = ast_depth(formula)
assert (depth is None) is expected, depth
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
assert any('complexity limit' in warning for warning in r.warnings)
print('survived')
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr


@pytest.mark.parametrize("case", ["nesting", "tokens", "source-bytes"])
def test_parser_complexity_limits_keep_formula_cache_and_clean_cells(case):
    script = f"""
import io, zipfile
import formualizer as fz
from openpyxl import Workbook
from linexcel.engine import boot_engine, is_too_deep
case = {case!r}
if case == "nesting":
    formula = "=" + "(" * 73 + "1" + ")" * 73
elif case == "tokens":
    formula = "=SUM(" + ",".join(["1"] * 9000) + ")"
else:
    formula = '="' + "é" * 32768 + '"'
try:
    fz.parse(formula)
except Exception as error:
    parser_limited = any(text in str(error) for text in (
        "AST height limit exceeded", "Formula nesting too deep",
        "node limit exceeded", "Formula token limit exceeded",
        "Formula source byte limit exceeded"))
else:
    parser_limited = False
assert is_too_deep(formula) is parser_limited
w = Workbook(); w.active.title = "S"
w.active["A1"] = formula; w.active["B1"] = "=2+3"
buffer = io.BytesIO(); w.save(buffer); cached = io.BytesIO()
with zipfile.ZipFile(buffer) as src, zipfile.ZipFile(cached, "w") as dst:
    for item in src.infolist():
        content = src.read(item.filename)
        if item.filename == "xl/worksheets/sheet1.xml":
            from xml.etree import ElementTree as ET
            root = ET.fromstring(content)
            cell = root.find(".//{{*}}c[@r='A1']")
            # openpyxl truncates strings at Excel's 32,767-character cell
            # limit; write the hostile formula directly into the XML so this
            # case reaches Formualizer's independent 65,536-byte parser cap.
            cell.find("{{*}}f").text = formula[1:]
            cell.find("{{*}}v").text = "99"
            content = ET.tostring(root)
        dst.writestr(item, content)
session = boot_engine(cached.getvalue(), [])
key = ("S", 1, 1)
if parser_limited:
    assert key in session.quarantined, session.quarantined
    assert session.engine.get_value("S", 1, 1) == 99
else:
    assert key not in session.quarantined, session.quarantined
assert session.engine.get_value("S", 1, 2) == 5
print("survived")
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr
    assert "survived" in process.stdout


def test_large_flat_sum_and_700_term_chain_still_recalculate():
    script = """
import io
import formualizer as fz
from openpyxl import Workbook
from linexcel.engine import boot_engine
w = Workbook(); w.active.title = 'S'
w.active['A1'] = '=SUM(' + ','.join(['1'] * 3000) + ')'
chain = '=' + '+'.join(['1'] * 700)
try:
    fz.parse(chain)
except Exception as error:
    parser_limited = 'AST height limit exceeded' in str(error)
else:
    parser_limited = False
w.active['B1'] = chain
output=io.BytesIO(); w.save(output)
session=boot_engine(output.getvalue(), [])
assert session.engine.get_value('S',1,1) == 3000
key = ('S', 1, 2)
if parser_limited:
    assert key in session.quarantined, session.quarantined
    assert session.engine.get_value('S',1,2) == {'type': 'Error', 'kind': 'NImpl'}
else:
    assert key not in session.quarantined, session.quarantined
    assert session.engine.get_value('S',1,2) == 700
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr


@pytest.mark.parametrize("terms", [256, 257])
def test_native_parser_height_boundary_is_respected_before_import(terms):
    script = f"""
import io
import formualizer as fz
from openpyxl import Workbook
from linexcel.engine import boot_engine, is_too_deep
formula = "=" + "+".join(["1"] * {terms})
try:
    fz.parse(formula)
except Exception as error:
    parser_limited = "AST height limit exceeded" in str(error)
else:
    parser_limited = False
assert is_too_deep(formula) is parser_limited
w = Workbook(); w.active.title = "S"; w.active["A1"] = formula
output = io.BytesIO(); w.save(output)
session = boot_engine(output.getvalue(), [])
key = ("S", 1, 1)
if parser_limited:
    assert key in session.quarantined, session.quarantined
else:
    assert key not in session.quarantined, session.quarantined
    assert session.engine.get_value("S", 1, 1) == {terms}
print("survived")
"""
    process = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, timeout=30
    )
    assert process.returncode == 0, process.stdout + process.stderr
    assert "survived" in process.stdout


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
