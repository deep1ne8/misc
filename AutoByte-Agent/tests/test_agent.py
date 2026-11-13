"""Agent loop tests, including the --mock end-to-end path."""

from __future__ import annotations

import builtins
import json
import pathlib
from collections.abc import Iterator

import pytest

import agent


def test_parse_tool_no_block_returns_none() -> None:
    assert agent.parse_tool("Just restart the spooler.") == (None, "")


@pytest.mark.parametrize(
    "text,expected",
    [
        ('```tool\n{"tool": "disk"}\n```', ("disk", "")),
        ('x\n```tool\n{"tool": "services", "arg": "Spooler"}\n```', ("services", "Spooler")),
        ('```tool\n{"tool": "event_log", "arg": 5}\n```', ("event_log", "5")),
        ('```tool\n{"tool": "disk", "arg": null}\n```', ("disk", "")),
        ("```tool\n{not json}\n```", (None, "")),
        ('```tool\n["disk"]\n```', (None, "")),
        ('```tool\n{"tool": 7}\n```', (None, "")),
    ],
)
def test_parse_tool_variants(text: str, expected: tuple[str | None, str]) -> None:
    assert agent.parse_tool(text) == expected


class Script:
    def __init__(self, replies: list[str]) -> None:
        self.replies = iter(replies)

    def chat(self, messages: list[dict[str, str]]) -> str:
        return next(self.replies)


def _run(model: object, answers: Iterator[bool] | None = None) -> tuple[list[str], list[dict[str, str]]]:
    out: list[str] = []
    hist = [{"role": "system", "content": "s"}, {"role": "user", "content": "q"}]
    confirm = (lambda _p: next(answers)) if answers else (lambda _p: False)
    agent.handle_turn(model, hist, confirm, out.append)  # type: ignore[arg-type]
    return out, hist


def test_plain_answer_does_not_crash() -> None:
    out, hist = _run(Script(["Hello there."]))
    assert any("Hello there." in o for o in out)
    assert hist[-1] == {"role": "assistant", "content": "Hello there."}


def test_unknown_tool_reported_not_crash() -> None:
    out, hist = _run(Script(['```tool\n{"tool": "nuke"}\n```', "ok"]))
    assert any("[error]" in o for o in out)
    assert hist[-2]["content"].startswith("Tool error")


def test_refused_command_never_prompts(runner: object) -> None:
    def no_prompt(_p: str) -> bool:
        raise AssertionError("must not prompt for a deny-listed command")

    out: list[str] = []
    hist = [{"role": "system", "content": "s"}, {"role": "user", "content": "wipe"}]
    agent.handle_turn(
        Script(['```tool\n{"tool": "run_command", "arg": "Format-Volume C"}\n```', "understood"]),
        hist,
        no_prompt,
        out.append,
    )
    assert any("[refused]" in o for o in out)
    assert runner.calls == []  # type: ignore[attr-defined]


def test_tool_rounds_are_bounded(runner: object) -> None:
    model = Script(['```tool\n{"tool": "disk"}\n```'] * (agent.MAX_TOOL_ROUNDS + 2))
    out, _ = _run(model)
    assert len(runner.calls) == agent.MAX_TOOL_ROUNDS  # type: ignore[attr-defined]
    assert "maximum number of tool calls" in out[-1]


def test_load_cfg_prefers_env_key(tmp_path: pathlib.Path, monkeypatch: pytest.MonkeyPatch) -> None:
    p = tmp_path / "c.json"
    p.write_text(json.dumps({"mode": "api", "api_key": "from-file"}))
    monkeypatch.setenv("OPENAI_API_KEY", "from-env")
    assert agent.load_cfg(p)["api_key"] == "from-env"


def test_mock_end_to_end(
    monkeypatch: pytest.MonkeyPatch, runner: object, capsys: pytest.CaptureFixture[str]
) -> None:
    inputs = iter(["hello", "check my disk", "wipe it", "my pc is slow", "n", "is a restart pending", "quit"])
    monkeypatch.setattr(builtins, "input", lambda _p="": next(inputs))
    assert agent.main(["--mock"]) == 0
    out = capsys.readouterr().out
    assert "mock responder" in out
    assert "[refused]" in out and "deny-list" in out  # wipe
    assert "Cancelled." in out  # declined CBSRepair
    ran = [c[-1] for c in runner.calls]  # type: ignore[attr-defined]
    assert any("Win32_LogicalDisk" in c for c in ran)  # disk
    assert any(c.endswith("CheckIfSystemNeedsToRestart.ps1") for c in ran)  # real manifest diagnostic
    assert not any("CBSRepair" in c for c in ran)
    assert not any("Format-Volume" in c for c in ran)
