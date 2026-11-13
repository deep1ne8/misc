"""AutoByte Agent - local Windows troubleshooting AI.

Loop per user turn (bounded to MAX_TOOL_ROUNDS tool calls):
  1. The model answers in plain English, or emits one fenced tool call:
       ```tool
       {"tool": "disk"}
       ```
  2. Read-only tools run automatically; mutating tools show exactly what will run
     and wait for an explicit "y". Destructive commands are refused outright.
  3. The tool result is fed back and the model explains it.

Models (config.json, created from config.example.json):
  - mode "local": llama-cpp-python loads ``model_path`` (a GGUF).
  - mode "api":   OpenAI-compatible endpoint. Key from OPENAI_API_KEY (preferred).
  - --mock:       canned responder for testing without a model.
"""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import re
import sys
from collections.abc import Callable
from typing import Any, Protocol

import tools

HERE = pathlib.Path(__file__).resolve().parent
DEFAULT_CFG = HERE / "config.json"
EXAMPLE_CFG = HERE / "config.example.json"
MAX_TOOL_ROUNDS = 4
MAX_HISTORY = 24  # messages kept after the system prompt
RESULT_TO_MODEL_CHARS = 6000

SYS = """You are AutoByte, a careful Windows IT troubleshooting assistant.
You help a technician or user diagnose, fix, repair, and update a Windows PC,
and answer Windows questions in plain English.

Rules:
- Prefer READ-ONLY tools first: system_info, disk, event_log [count], network,
  services <name>, run_diagnostic <exact script name>.
- Mutating tools (run_fix <exact script name>, run_command <powershell>, install_updates)
  need operator approval. State what it will do and the risk before calling one.
- Destructive commands (disk format, boot config, recursive deletes of system paths,
  registry hive deletes, iex/download-and-run) are always refused; do not propose them.
- To use a tool, output EXACTLY one fenced block and nothing after it:
  ```tool
  {"tool": "services", "arg": "Spooler"}
  ```
- If you don't need a tool, just answer in plain English. Be concise and practical.
"""

TOOL_RE = re.compile(r"```tool\s*(\{.*?\})\s*```", re.DOTALL)


class ChatModel(Protocol):
    def chat(self, messages: list[dict[str, str]]) -> str: ...


# ---- config --------------------------------------------------------------------
def load_cfg(path: pathlib.Path = DEFAULT_CFG) -> dict[str, Any]:
    cfg: dict[str, Any] = {
        "mode": "local",
        "model_path": "",
        "api_base": "https://api.openai.com/v1",
        "api_key": "",
        "model": "gpt-4o-mini",
    }
    for p in (EXAMPLE_CFG, path):
        if p.exists():
            try:
                cfg.update(json.loads(p.read_text(encoding="utf-8")))
            except ValueError as e:
                raise SystemExit(f"Invalid JSON in {p}: {e}") from e
    # Environment wins so real keys never need to live in a file.
    cfg["api_key"] = os.environ.get("OPENAI_API_KEY") or cfg.get("api_key", "")
    cfg["api_base"] = os.environ.get("OPENAI_BASE_URL") or cfg["api_base"]
    return cfg


# ---- model backends --------------------------------------------------------------
class LocalModel:
    def __init__(self, model_path: str) -> None:
        if not model_path or not pathlib.Path(model_path).is_file():
            raise SystemExit(
                f"No local model found at {model_path!r}. Run setup_model.py, "
                "set config.json model_path, or use --mock / api mode."
            )
        from llama_cpp import Llama

        self.llm = Llama(model_path=str(model_path), n_ctx=8192, n_threads=os.cpu_count() or 8, verbose=False)

    def chat(self, messages: list[dict[str, str]]) -> str:
        out: Any = self.llm.create_chat_completion(messages=messages, temperature=0.3, max_tokens=1024)
        return str(out["choices"][0]["message"]["content"] or "")


class ApiModel:
    def __init__(self, base: str, key: str, model: str) -> None:
        try:
            from openai import OpenAI
        except ImportError as e:
            raise SystemExit("openai package not installed: pip install -r requirements-api.txt") from e
        if not key:
            raise SystemExit("api mode needs OPENAI_API_KEY set in the environment.")
        self.client = OpenAI(base_url=base, api_key=key)
        self.model = model

    def chat(self, messages: list[dict[str, str]]) -> str:
        r = self.client.chat.completions.create(
            model=self.model,
            messages=messages,  # type: ignore[arg-type]
            temperature=0.3,
        )
        return r.choices[0].message.content or ""


class MockModel:
    """Canned responder: exercises the tool loop and safety gate without a model."""

    def chat(self, messages: list[dict[str, str]]) -> str:
        last = messages[-1]["content"]
        low = last.lower()
        if last.startswith(("Tool result", "Tool refused", "Operator declined", "Tool error")):
            return "Done. Summary: see the result above."
        if "disk" in low:
            return 'Let me check disk space.\n```tool\n{"tool": "disk"}\n```'
        if "restart" in low:
            return '```tool\n{"tool": "run_diagnostic", "arg": "CheckIfSystemNeedsToRestart"}\n```'
        if "wipe" in low:
            return '```tool\n{"tool": "run_command", "arg": "Format-Volume -DriveLetter C"}\n```'
        if "slow" in low or "update" in low:
            return (
                "A component-store repair is a good first step.\n"
                '```tool\n{"tool": "run_fix", "arg": "CBSRepair"}\n```'
            )
        return (
            "I'm a mock responder. Ask about 'disk', 'restart', 'slow PC' or 'wipe' "
            "to see the tool and safety flow."
        )


# ---- tool parsing ------------------------------------------------------------------
def parse_tool(text: str) -> tuple[str | None, str]:
    """Return (tool, arg). tool is None when no valid tool block is present."""
    m = TOOL_RE.search(text or "")
    if not m:
        return None, ""
    try:
        d = json.loads(m.group(1))
    except ValueError:
        return None, ""
    if not isinstance(d, dict) or not isinstance(d.get("tool"), str):
        return None, ""
    arg = d.get("arg", "")
    return d["tool"], "" if arg is None else str(arg)


# ---- one user turn -------------------------------------------------------------------
def handle_turn(
    model: ChatModel,
    history: list[dict[str, str]],
    confirm: Callable[[str], bool],
    emit: Callable[[str], None] = print,
) -> None:
    for _ in range(MAX_TOOL_ROUNDS):
        reply = model.chat(history)
        history.append({"role": "assistant", "content": reply})
        tool, arg = parse_tool(reply)
        if tool is None:
            emit("\nAutoByte> " + reply.strip() + "\n")
            return
        prose = TOOL_RE.split(reply)[0].strip()
        if prose:
            emit("\nAutoByte> " + prose)

        res = tools.call(tool, arg)
        status = res.get("status")
        if status == "need_confirm":
            emit(f"\n  Proposed action ({tool}):\n    {res.get('preview', '')}")
            res = tools.execute(tool, arg, approved=confirm("  Approve? [y/N] "))
            status = res.get("status")
            if status == "refused":
                emit("  Cancelled.\n")
                history.append({"role": "user", "content": "Operator declined the action."})
                continue

        out = str(res.get("result", ""))
        if status == "ok":
            emit("\n--- result ---\n" + out[:2000] + "\n")
            history.append({"role": "user", "content": "Tool result:\n" + out[:RESULT_TO_MODEL_CHARS]})
        elif status == "refused":
            emit(f"\n  [refused] {out}\n")
            history.append({"role": "user", "content": f"Tool refused: {out}"})
        else:
            emit(f"\n  [error] {out}\n")
            history.append({"role": "user", "content": f"Tool error: {out}"})
    emit("\nAutoByte> (stopped after the maximum number of tool calls for one question)\n")


def _trim(history: list[dict[str, str]]) -> None:
    if len(history) > MAX_HISTORY + 1:
        del history[1 : len(history) - MAX_HISTORY]


def build_model(args: argparse.Namespace, cfg: dict[str, Any]) -> ChatModel:
    if args.mock:
        print("[mode] mock responder")
        return MockModel()
    if cfg.get("mode") == "api":
        print("[mode] api:", cfg.get("model"))
        return ApiModel(cfg["api_base"], cfg.get("api_key", ""), cfg["model"])
    print("[mode] local GGUF")
    return LocalModel(cfg.get("model_path", ""))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="AutoByte local Windows troubleshooting agent")
    ap.add_argument("--mock", action="store_true", help="use canned responder (no model)")
    ap.add_argument("--config", default=str(DEFAULT_CFG))
    args = ap.parse_args(argv)

    model = build_model(args, load_cfg(pathlib.Path(args.config)))
    if not tools.load_manifest():
        print("[warn] scripts.manifest.json not found: script tools disabled.", file=sys.stderr)
    print("AutoByte Agent ready. Type a Windows issue or question. 'quit' to exit.\n")
    history: list[dict[str, str]] = [{"role": "system", "content": SYS}]

    def confirm(prompt: str) -> bool:
        try:
            return input(prompt).strip().lower() in ("y", "yes")
        except (EOFError, KeyboardInterrupt):
            return False

    while True:
        try:
            q = input("You> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nbye")
            return 0
        if not q:
            continue
        if q.lower() in ("quit", "exit"):
            return 0
        mark = len(history)
        history.append({"role": "user", "content": q})
        try:
            handle_turn(model, history, confirm)
        except Exception as e:  # keep the session alive on backend/network errors
            print(f"\n  [error] {type(e).__name__}: {e}\n")
            del history[mark:]
        _trim(history)


if __name__ == "__main__":
    sys.exit(main())
