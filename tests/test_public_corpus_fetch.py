"""Public downloads stay pinned, bounded and static; these tests use no network."""

import hashlib
import importlib.util
from pathlib import Path
from urllib.request import Request
from zipfile import ZipFile

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "scripts/fetch_public_corpus.py"
SPEC = importlib.util.spec_from_file_location("public_corpus_fetch", SCRIPT)
fetch = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(fetch)


def source(data=b"source bytes", **changes):
    return {
        "filename": "source.xla",
        "source_url": "https://downloads.example.org/source.xla",
        "sha256": hashlib.sha256(data).hexdigest(),
        "max_bytes": 1024,
        **changes,
    }


def test_existing_file_must_match_pin_and_is_never_overwritten(tmp_path):
    path = tmp_path / "source.xla"
    path.write_bytes(b"locally modified")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        fetch.acquire(source(), tmp_path, offline=True)
    assert path.read_bytes() == b"locally modified"


def test_verified_existing_file_works_without_network(tmp_path, monkeypatch):
    data = b"source bytes"
    (tmp_path / "source.xla").write_bytes(data)
    monkeypatch.setattr(
        fetch, "build_opener", lambda *_: pytest.fail("Network attempted")
    )
    receipt = fetch.acquire(source(data), tmp_path, offline=True)
    assert receipt["status"] == "reused_verified"
    assert receipt["sha256"] == hashlib.sha256(data).hexdigest()


@pytest.mark.parametrize(
    "filename", ["../escape.xla", "C:\\escape.xla", "sub/file.xla"]
)
def test_source_filename_cannot_escape_output(filename):
    with pytest.raises(ValueError, match="basename"):
        fetch.validate_source(source(filename=filename))


@pytest.mark.parametrize(
    "url",
    [
        "http://downloads.example.org/source.xla",
        "https://user:password@downloads.example.org/source.xla",
    ],
)
def test_sources_require_https_without_credentials(url):
    with pytest.raises(ValueError, match="HTTPS"):
        fetch.validate_source(source(source_url=url))


def test_redirect_must_stay_on_allowed_https_hosts():
    policy = fetch.CorpusRedirects(
        source(redirect_host_suffixes=["dl.sourceforge.net"])
    )
    request = Request("https://downloads.example.org/source.xla")
    allowed = policy.redirect_request(
        request,
        None,
        302,
        "Found",
        {},
        "https://master.dl.sourceforge.net/project/source.xla",
    )
    assert allowed.full_url.startswith("https://master.dl.sourceforge.net/")
    for url in (
        "http://downloads.example.org/file",
        "https://dl.sourceforge.net.attacker.org/file",
        "https://unrelated.example.org/file",
    ):
        with pytest.raises(ValueError, match="redirect"):
            policy.redirect_request(request, None, 302, "Found", {}, url)


class Response:
    def __init__(self, data, *, length=None):
        self.data = data
        self.headers = {} if length is None else {"Content-Length": str(length)}
        self.url = "https://downloads.example.org/source.xla"

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False

    def read(self, count):
        return self.data[:count]


@pytest.mark.parametrize("length", [None, 2048])
def test_oversize_download_is_rejected_without_creating_source(
    tmp_path, monkeypatch, length
):
    class Opener:
        def open(self, *_args, **_kwargs):
            return Response(b"x" * 2048, length=length)

    monkeypatch.setattr(fetch, "build_opener", lambda *_: Opener())
    with pytest.raises(ValueError, match="size bound"):
        fetch.acquire(source(), tmp_path, offline=False)
    assert not (tmp_path / "source.xla").exists()


def test_file_created_during_download_is_not_replaced(tmp_path, monkeypatch):
    class Opener:
        def open(self, *_args, **_kwargs):
            (tmp_path / "source.xla").write_bytes(b"concurrent user content")
            return Response(b"source bytes")

    monkeypatch.setattr(fetch, "build_opener", lambda *_: Opener())
    with pytest.raises(ValueError, match="changed during acquisition"):
        fetch.acquire(source(), tmp_path, offline=False)
    assert (tmp_path / "source.xla").read_bytes() == b"concurrent user content"


def test_html_disguised_as_addin_is_rejected(tmp_path):
    path = tmp_path / "source.xla"
    path.write_bytes(b"<html>not an add-in</html>")
    with pytest.raises(ValueError, match="OLE"):
        fetch.inspect_container(path)


def test_zip_member_cannot_escape_static_inspection(tmp_path):
    path = tmp_path / "source.xlsx"
    with ZipFile(path, "w") as archive:
        archive.writestr("../payload", "not an Excel package")
    with pytest.raises(ValueError, match="Unsafe"):
        fetch.inspect_container(path)
