from __future__ import annotations

import hashlib
import json
import pathlib
import sys
from typing import Any

import pytest

AGENT_DIR = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(AGENT_DIR))

import tools  # noqa: E402


class FakeRunner:
    """Records every subprocess invocation instead of running PowerShell."""

    def __init__(self) -> None:
        self.calls: list[list[str]] = []

    def __call__(self, args: list[str], timeout: int) -> tuple[int, str]:
        self.calls.append(args)
        return 0, "fake-output"


@pytest.fixture()
def runner(monkeypatch: pytest.MonkeyPatch) -> FakeRunner:
    r = FakeRunner()
    monkeypatch.setattr(tools, "_run", r)
    return r


@pytest.fixture(autouse=True)
def audit_dir(tmp_path: pathlib.Path, monkeypatch: pytest.MonkeyPatch) -> pathlib.Path:
    d = tmp_path / "logs"
    monkeypatch.setenv("AUTOBYTE_LOG_DIR", str(d))
    return d


def _entry(name: str, body: bytes, **flags: Any) -> tuple[dict[str, Any], bytes]:
    e = {
        "name": name,
        "path": f"Scripts/{name}.ps1",
        "sha256": hashlib.sha256(body.replace(b"\r\n", b"\n")).hexdigest(),
        "platform": "windows",
        "mutating": True,
        "interactive": False,
        "unsupported": False,
        "agentDiagnostic": False,
    }
    e.update(flags)
    return e, body


@pytest.fixture()
def fake_repo(tmp_path: pathlib.Path, monkeypatch: pytest.MonkeyPatch) -> pathlib.Path:
    root = tmp_path / "repo"
    scripts = root / "Scripts"
    scripts.mkdir(parents=True)
    specs = [
        _entry("CheckDriveSpace", b"Get-PSDrive C\r\n", mutating=False, agentDiagnostic=True),
        _entry("CBSRepair", b"DISM /Online /Cleanup-Image /RestoreHealth\n"),
        _entry("UpdateWin11-BypassTpm", b"# bypass\n", unsupported=True),
        _entry("AskUser", b"Read-Host 'x'\n", interactive=True),
        _entry("Tampered", b"original\n", mutating=False, agentDiagnostic=True),
    ]
    entries = []
    for e, body in specs:
        (root / e["path"]).write_bytes(body)
        entries.append(e)
    (scripts / "Tampered.ps1").write_bytes(b"Format-Volume C\n")
    (root / "scripts.manifest.json").write_text(json.dumps({"scripts": entries}))
    monkeypatch.setattr(tools, "REPO_ROOT", root)
    monkeypatch.setattr(tools, "SCRIPTS_DIR", scripts)
    monkeypatch.setattr(tools, "MANIFEST_PATH", root / "scripts.manifest.json")
    return root
