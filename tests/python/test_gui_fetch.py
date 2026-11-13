"""AutoByteGUI download path: only manifest-listed, hash-matching scripts are accepted."""

from __future__ import annotations

import hashlib
import json
import pathlib
import sys
import types
from typing import Any

import pytest

# The download/verify logic doesn't need a display; stub tkinter when it's absent.
if "tkinter" not in sys.modules:
    try:
        import tkinter  # noqa: F401
    except ImportError:
        tk_stub = types.ModuleType("tkinter")
        st_stub = types.ModuleType("tkinter.scrolledtext")
        tk_stub.Tk = tk_stub.Button = tk_stub.Frame = tk_stub.Label = tk_stub.Entry = object  # type: ignore[attr-defined]
        tk_stub.END = "end"  # type: ignore[attr-defined]
        st_stub.ScrolledText = object  # type: ignore[attr-defined]
        sys.modules["tkinter"] = tk_stub
        sys.modules["tkinter.scrolledtext"] = st_stub

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import AutoByteGUI as gui  # noqa: E402


class Resp:
    def __init__(self, content: bytes) -> None:
        self.content = content

    def raise_for_status(self) -> None:
        pass

    def json(self) -> Any:
        return json.loads(self.content)


def _manifest(body: bytes, **extra: Any) -> dict[str, dict[str, Any]]:
    e = {"name": "Good", "path": "Scripts/Good.ps1", "sha256": hashlib.sha256(body).hexdigest()}
    e.update(extra)
    return {"good": e}


def test_verified_download(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(gui.requests, "get", lambda url, timeout: Resp(b"ok\r\n"))
    assert gui.fetch_verified(_manifest(b"ok\n"), "Good") == b"ok\r\n"  # CRLF normalised


def test_tampered_download_rejected(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(gui.requests, "get", lambda url, timeout: Resp(b"EVIL"))
    with pytest.raises(ValueError, match="mismatch"):
        gui.fetch_verified(_manifest(b"ok\n"), "Good")


def test_unknown_and_unsupported_rejected() -> None:
    with pytest.raises(ValueError, match="not in the manifest"):
        gui.fetch_verified(_manifest(b"x"), "Other")
    with pytest.raises(ValueError, match="unsupported"):
        gui.fetch_verified(_manifest(b"x", unsupported=True), "Good")


def test_gui_buttons_exist_in_committed_manifest() -> None:
    names = {s["name"] for s in json.loads((ROOT / "scripts.manifest.json").read_text())["scripts"]}
    missing = [n for _, n in gui.GUI_SCRIPTS if n not in names]
    assert not missing


def test_default_ref_is_pinned() -> None:
    assert gui.PINNED_RE.match("v1.0.0") and not gui.PINNED_RE.match("main")
