"""
AutoByte Agent — local Windows troubleshooting AI.

Loop:
  1. You type a question in plain English.
  2. The model reasons and may emit a TOOL call inside a fenced block:
       ```tool
       {"tool": "disk"}            # read-only -> runs automatically
       {"tool": "run_fix", "arg": "CBSRepair"}   # mutating -> asks Y/n
       ```
  3. Tool results are fed back; the model explains + advises.
  4. When no tool is needed, the model answers in plain English.

Models (config.json):
  - mode "local":  llama-cpp-python loads config["model_path"] (a GGUF).
  - mode "api":    OpenAI-compatible endpoint (needs OPENAI_API_KEY / config key).
  - --mock:        canned responder so the framework can be tested without a model.

Safety: mutating tools never run without an explicit Y from the operator.
"""
import json, sys, pathlib, re, os

HERE = pathlib.Path(__file__).resolve().parent
DEFAULT_CFG = HERE / "config.json"

SYS = """You are AutoByte, a careful Windows IT troubleshooting assistant.
You help a technician or user diagnose, fix, repair, and update a Windows PC,
and answer any Windows question in plain English.

Rules:
- Prefer READ-ONLY diagnostics first (system_info, disk, event_log, network, services, run_diagnostic).
- Only suggest mutating actions (run_fix, run_command, install_updates) when clearly needed,
  and always state what it will do and the risk before calling it.
- To use a tool, output EXACTLY one fenced block:
  ```tool
  {"tool": "disk"}
  ```
  or with an argument:
  ```tool
  {"tool": "services", "arg": "Spooler"}
  ```
- If you don't need a tool, just answer in plain English.
- Keep answers concise and practical. No fluff.
"""


def load_cfg(path=DEFAULT_CFG):
    if path.exists():
        return json.loads(path.read_text())
    return {"mode": "local", "model_path": "", "api_base": "https://api.openai.com/v1", "api_key": os.environ.get("OPENAI_API_KEY", ""), "model": "gpt-4o-mini"}


# ---- model backends --------------------------------------------------------
class LocalModel:
    def __init__(self, model_path):
        if not model_path or not pathlib.Path(model_path).exists():
            raise SystemExit("No local model found at %s. Set config.json model_path or use --mock / api mode." % model_path)
        from llama_cpp import Llama
        self.llm = Llama(model_path=str(model_path), n_ctx=8192, n_threads=os.cpu_count() or 8)

    def chat(self, messages):
        out = self.llm.create_chat_completion(messages=messages, temperature=0.3, max_tokens=1024)
        return out["choices"][0]["message"]["content"]


class ApiModel:
    def __init__(self, base, key, model):
        try:
            from openai import OpenAI
        except ImportError:
            raise SystemExit("openai package not installed. pip install openai")
        self.client = OpenAI(base_url=base, api_key=key or "EMPTY")
        self.model = model

    def chat(self, messages):
        r = self.client.chat.completions.create(model=self.model, messages=messages, temperature=0.3)
        return r.choices[0].message.content


class MockModel:
    """Canned responder: demonstrates the tool loop without a real model."""
    def chat(self, messages):
        last = messages[-1]["content"]
        if "disk" in last.lower():
            return 'Let me check disk space.\n```tool\n{"tool": "disk"}\n```'
        if "slow" in last.lower() or "update" in last.lower():
            return ('A good first step is a component-store repair, then Windows Update.\n'
                    '```tool\n{"tool": "run_fix", "arg": "CBSRepair"}\n```')
        return "I'm a mock responder. I can confirm the tool-calling and confirmation flow works. Ask about 'disk' or 'slow PC' to see tools fire."


# ---- tool parsing ----------------------------------------------------------
TOOL_RE = re.compile(r"```tool\s*(\{.*?\})\s*```", re.DOTALL)

def parse_tool(text):
    m = TOOL_RE.search(text)
    if not m:
        return None
    try:
        d = json.loads(m.group(1))
        return d.get("tool"), d.get("arg", "")
    except Exception:
        return None, None


# ---- main loop --------------------------------------------------------------
def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--mock", action="store_true", help="use canned responder (no model download)")
    ap.add_argument("--config", default=str(DEFAULT_CFG))
    args = ap.parse_args()

    cfg = load_cfg(pathlib.Path(args.config))
    if args.mock:
        model = MockModel()
        print("[mode] mock responder")
    elif cfg.get("mode") == "api":
        model = ApiModel(cfg["api_base"], cfg.get("api_key", ""), cfg["model"])
        print("[mode] api:", cfg.get("model"))
    else:
        model = LocalModel(cfg.get("model_path", ""))
        print("[mode] local GGUF")

    import tools
    print("AutoByte Agent ready. Type a Windows issue or question. 'quit' to exit.\n")
    history = [{"role": "system", "content": SYS}]

    while True:
        try:
            q = input("You> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nbye"); break
        if not q:
            continue
        if q.lower() in ("quit", "exit"):
            break
        history.append({"role": "user", "content": q})

        # one tool-call round (keeps it simple + safe)
        reply = model.chat(history)
        tool, arg = parse_tool(reply)
        if tool:
            res = tools.call(tool, arg)
            if isinstance(res, dict) and res.get("need_confirm"):
                print(f"\n  Proposed action: {tool} {('('+arg+')') if arg else ''}")
                ans = input("  Approve? [y/N] ").strip().lower()
                if ans == "y":
                    out = tools.DISPATCH[tool](arg)
                    print("\n--- result ---\n" + out[:2000] + "\n")
                    history.append({"role": "assistant", "content": reply})
                    history.append({"role": "user", "content": "Tool result:\n" + out[:3000]})
                    follow = model.chat(history)
                    print("\nAutoByte> " + follow.strip() + "\n")
                    history.append({"role": "assistant", "content": follow})
                else:
                    print("  Cancelled.\n")
                    history.append({"role": "assistant", "content": reply})
                    history.append({"role": "user", "content": "Operator declined the action."})
            else:
                out = res.get("result", str(res))
                print("\n--- result ---\n" + out[:2000] + "\n")
                history.append({"role": "assistant", "content": reply})
                history.append({"role": "user", "content": "Tool result:\n" + out[:3000]})
                follow = model.chat(history)
                print("\nAutoByte> " + follow.strip() + "\n")
                history.append({"role": "assistant", "content": follow})
        else:
            print("\nAutoByte> " + reply.strip() + "\n")
            history.append({"role": "assistant", "content": reply})


if __name__ == "__main__":
    main()
