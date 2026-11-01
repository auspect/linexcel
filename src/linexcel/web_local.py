"""Persistent local identity and one-use browser access links.

Only the standalone launcher enables this profile. Embedded applications keep
their host-provided identity resolver and never adopt local workspaces.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import secrets
import sys
import time
from pathlib import Path
from urllib.parse import urlsplit

TOKEN_SECONDS = 600
COOKIE_SECONDS = 400 * 86400


def _process_exists(pid: object) -> bool:
    if not isinstance(pid, int) or isinstance(pid, bool) or pid <= 0:
        return False
    if sys.platform != "win32":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    # Windows os.kill(pid, 0) uses TerminateProcess; query a handle instead.
    import ctypes
    from ctypes import wintypes

    api = ctypes.WinDLL("kernel32", use_last_error=True)
    api.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    api.OpenProcess.restype = wintypes.HANDLE
    api.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    api.GetExitCodeProcess.restype = wintypes.BOOL
    api.CloseHandle.argtypes = [wintypes.HANDLE]
    api.CloseHandle.restype = wintypes.BOOL
    handle = api.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
    if not handle:
        return False
    try:
        code = wintypes.DWORD()
        return (
            bool(api.GetExitCodeProcess(handle, ctypes.byref(code)))
            and code.value == 259
        )
    finally:
        api.CloseHandle(handle)


def _private_write(path: Path, value: dict) -> None:
    temporary = path.with_name(path.name + "." + secrets.token_hex(8) + ".tmp")
    descriptor = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(value, stream)
        temporary.replace(path)
        path.chmod(0o600)
    finally:
        temporary.unlink(missing_ok=True)


def loopback_request(environ: dict) -> bool:
    import ipaddress

    try:
        peer = ipaddress.ip_address(environ.get("REMOTE_ADDR", ""))
        host = urlsplit("//" + environ.get("HTTP_HOST", "")).hostname
        return peer.is_loopback and host in {"localhost", "127.0.0.1", "::1"}
    except ValueError:
        return False


class LocalProfile:
    def __init__(self, root: Path):
        self.root = root
        self.directory = root / ".access"
        self.directory.mkdir(mode=0o700, exist_ok=True)
        self.directory.chmod(0o700)
        path = self.directory / "profile.json"
        if path.exists():
            value = json.loads(path.read_text("utf-8"))
            identity = value.get("identity")
            if value.get("version") != 1 or not isinstance(identity, str):
                raise ValueError("Invalid local profile; restore its private backup")
            if not re.fullmatch(r"[a-f0-9]{64}", identity):
                raise ValueError("Invalid local identity")
            self.identity = identity
        else:
            self.identity = secrets.token_hex(32)
            _private_write(path, {"version": 1, "identity": self.identity})
        self.owner = hashlib.sha256(self.identity.encode()).hexdigest()
        self._publication = None

    def publish(self, url: str) -> str:
        publication = secrets.token_hex(16)
        _private_write(
            self.directory / "server.json",
            {
                "url": url,
                "pid": os.getpid(),
                "startedAt": time.time(),
                "publication": publication,
            },
        )
        self._publication = publication
        return session_link(self.root)

    def consume(self, token: object) -> bool:
        if not isinstance(token, str) or not re.fullmatch(r"[a-f0-9]{64}", token):
            return False
        path = self.directory / (hashlib.sha256(token.encode()).hexdigest() + ".json")
        # Rename claims the token atomically across concurrent requests.
        claimed = path.with_name(path.name + ".used")
        try:
            path.rename(claimed)
        except FileNotFoundError:
            return False
        try:
            value = json.loads(claimed.read_text("utf-8"))
            return (
                value.get("owner") == self.owner
                and isinstance(value.get("expiresAt"), (int, float))
                and time.time() < value["expiresAt"]
            )
        except (OSError, ValueError, TypeError):
            return False
        finally:
            claimed.unlink(missing_ok=True)

    def close(self) -> None:
        if self._publication is None:
            return
        path = self.directory / "server.json"
        try:
            server = json.loads(path.read_text("utf-8"))
        except (FileNotFoundError, ValueError):
            self._publication = None
            return
        if isinstance(server, dict) and server.get("publication") == self._publication:
            path.unlink(missing_ok=True)
        self._publication = None


def session_link(root: Path) -> str:
    """Create a short-lived link using OS access to the private data directory."""
    root = Path(root).resolve()
    directory = root / ".access"
    try:
        server = json.loads((directory / "server.json").read_text("utf-8"))
        profile = json.loads((directory / "profile.json").read_text("utf-8"))
        url = server["url"]
        parsed = urlsplit(url)
        if parsed.scheme != "http" or parsed.hostname != "127.0.0.1":
            raise ValueError("Invalid local launcher address")
        if not _process_exists(server["pid"]):
            raise ValueError("Local launcher is no longer running")
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise ValueError("Start linexcel serve with this data directory first") from exc
    token = secrets.token_hex(32)
    _private_write(
        directory / (hashlib.sha256(token.encode()).hexdigest() + ".json"),
        {
            "owner": hashlib.sha256(profile["identity"].encode()).hexdigest(),
            "expiresAt": time.time() + TOKEN_SECONDS,
        },
    )
    return url + "#session=" + token
