"""Exercise the launcher with real loopback ports and browser failure recovery."""

import errno
import socket
import threading
import webbrowser
from unittest.mock import Mock, call

import pytest

from linexcel.cli import _build_parser, main


@pytest.mark.parametrize("port", ["-1", "65536", "not-a-port"])
def test_serve_rejects_invalid_ports_before_startup(port, capsys, monkeypatch):
    def unexpected_start(*args, **kwargs):
        pytest.fail("an invalid port must not initialize project storage")

    monkeypatch.setattr("linexcel.web.create_app", unexpected_start)
    with pytest.raises(SystemExit) as exc:
        main(["serve", "--port", port])
    assert exc.value.code == 2
    assert "--port" in capsys.readouterr().err


def test_serve_reports_occupied_port_and_releases_app(capsys, monkeypatch):
    app = Mock()
    monkeypatch.setattr("linexcel.web.create_app", lambda *args, **kwargs: app)
    with socket.socket() as occupied:
        occupied.bind(("127.0.0.1", 0))
        occupied.listen()
        port = occupied.getsockname()[1]
        assert main(["serve", "--port", str(port)]) == 2
    error = capsys.readouterr().err
    assert f"Port {port} is already in use" in error
    assert "--port 0" in error
    app.close.assert_called_once()


def test_serve_binds_exclusively_when_windows_socket_option_is_available(
    capsys, monkeypatch
):
    app = Mock()
    listener = Mock()
    listener.bind.side_effect = OSError(errno.EADDRINUSE, "Address already in use")
    exclusive_option = getattr(socket, "SO_EXCLUSIVEADDRUSE", -5)
    monkeypatch.setattr(socket, "SO_EXCLUSIVEADDRUSE", exclusive_option, raising=False)
    monkeypatch.setattr("linexcel.web.create_app", lambda *args, **kwargs: app)

    def make_server(host, port, app, *, server_class, handler_class):
        server = server_class.__new__(server_class)
        server.server_address = (host, port)
        server.socket = listener
        server.server_bind()
        pytest.fail("binding to an occupied port must fail")

    monkeypatch.setattr("wsgiref.simple_server.make_server", make_server)
    assert main(["serve", "--port", "8765"]) == 2
    assert "Port 8765 is already in use" in capsys.readouterr().err
    assert listener.method_calls == [
        call.setsockopt(socket.SOL_SOCKET, exclusive_option, 1),
        call.bind(("127.0.0.1", 8765)),
    ]
    app.close.assert_called_once()


@pytest.mark.parametrize("winerror", [None, 10013])
def test_serve_preserves_bind_permission_errors(winerror, capsys, monkeypatch):
    app = Mock()
    failure = PermissionError(errno.EACCES, "Binding is forbidden")
    if winerror is not None:
        failure.winerror = winerror
    monkeypatch.setattr("linexcel.web.create_app", lambda *args, **kwargs: app)
    monkeypatch.setattr("wsgiref.simple_server.make_server", Mock(side_effect=failure))
    assert main(["serve", "--port", "8765"]) == 2
    error = capsys.readouterr().err
    assert "Binding is forbidden" in error
    assert "already in use" not in error
    app.close.assert_called_once()


def test_serve_opens_actual_automatic_port_with_mount_prefix(capsys, monkeypatch):
    import linexcel.web

    app = Mock()
    opened = []
    finished = threading.Event()

    def open_browser(url):
        opened.append(url)
        finished.set()
        return True

    monkeypatch.setattr(linexcel.web, "create_app", lambda *args, **kwargs: app)
    monkeypatch.setattr("webbrowser.open", open_browser)
    monkeypatch.setattr(
        "wsgiref.simple_server.WSGIServer.serve_forever", lambda _: None
    )
    assert main(["serve", "--port", "0", "--mount-path", "/review", "--open"]) == 0
    assert finished.wait(timeout=2)
    output = capsys.readouterr().out.strip()
    assert output == f"Linexcel : {opened[0]}"
    assert opened[0].startswith("http://127.0.0.1:")
    assert opened[0].endswith("/review/")
    assert ":0/" not in opened[0]
    app.close.assert_called_once()


def test_serve_leaves_browser_closed_by_default():
    assert not _build_parser().parse_args(["serve"]).open_browser


@pytest.mark.parametrize("thinking", ["enabled", "disabled"])
def test_serve_selects_llama_template_thinking_independently(thinking):
    args = _build_parser().parse_args(
        ["serve", "--ai-chat-template-thinking", thinking]
    )
    assert args.ai_chat_template_thinking == thinking
    assert args.ai_thinking is None


@pytest.mark.parametrize("failure", [False, webbrowser.Error("no browser available")])
def test_browser_failure_keeps_manual_url_available(failure, capsys, monkeypatch):
    from linexcel.web import _open_serve_browser

    def open_browser(url):
        if isinstance(failure, Exception):
            raise failure
        return failure

    monkeypatch.setattr("webbrowser.open", open_browser)
    _open_serve_browser("http://127.0.0.1:12345/review/")
    assert "Open http://127.0.0.1:12345/review/" in capsys.readouterr().err
