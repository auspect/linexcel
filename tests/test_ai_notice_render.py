"""Exercise qualification excerpts through the viewer's actual JS renderer."""

import json
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest

from linexcel.aidoc import _qualify_documentation


class _Rendered(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tags = []
        self.code = []
        self.in_code = False

    def handle_starttag(self, tag, attrs):
        self.tags.append(tag)
        if tag == "code":
            self.in_code = True

    def handle_endtag(self, tag):
        if tag == "code":
            self.in_code = False

    def handle_data(self, data):
        if self.in_code:
            self.code.append(data)


@pytest.mark.parametrize("backticks", [False, True])
def test_qualification_uses_only_supported_viewer_code_spans(backticks):
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js required to execute the actual viewer renderer")
    quote = '=IF(A1,"**bold** <img src=x>&",B1)'
    if backticks:
        quote = quote.replace("**bold**", "`**bold**`")
    report = {
        "status": "qualified",
        "checks": [
            {"status": "not_source_supported", "quotation": quote},
        ],
    }
    markdown = _qualify_documentation("Original response.", report, "en")
    viewer = (Path(__file__).parents[1] / "src/linexcel/assets/viewer.html").read_text(
        encoding="utf-8"
    )
    functions = viewer.split("  function _md(src) {", 1)[1].split(
        "  // Mini syntax highlighting", 1
    )[0]
    script = (
        "function _md(src) {"
        + functions
        + (
            "\nprocess.stdout.write(_md(JSON.parse(require('fs').readFileSync(0,'utf8'))));"
        )
    )
    result = subprocess.run(
        [node, "-e", script],
        input=json.dumps(markdown),
        text=True,
        capture_output=True,
        timeout=10,
        check=True,
    )
    parsed = _Rendered()
    parsed.feed(result.stdout)
    assert "strong" not in parsed.tags
    assert "em" not in parsed.tags
    assert "img" not in parsed.tags
    if backticks:
        assert parsed.code == []
        assert "Quotation 1: excerpt omitted (Markdown delimiters)." in result.stdout
    else:
        assert "".join(parsed.code) == " " + quote + " "
    assert markdown.endswith("\n\nOriginal response.")
