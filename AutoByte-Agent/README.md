# AutoByte Agent

A **lightweight, local Windows AI troubleshooting agent**. It talks to you in
plain English, inspects your PC with safe PowerShell probes, explains what's
wrong, and — with your explicit approval — runs fixes from the bundled
AutoByte script library or its own repair commands.

- **Local & private**: runs a small GGUF model on your CPU (no data leaves the machine, no token limit).
- **Safe by design**: read-only diagnostics run automatically; anything that *changes* your system (repairs, updates, arbitrary commands) requires you to type `Y`.
- **Tool-using**: reasons over `system_info`, `disk`, `event_log`, `network`, `services`, and any of the 93 bundled scripts.

## Quick start

```powershell
cd AutoByte-Agent
pip install -r requirements.txt

# 1) get a local model (~4.7GB, one-time):
python setup_model.py

# 2) talk to it:
python agent.py
```

No model yet? Test the framework without a download:

```powershell
python agent.py --mock
```

## How it works

You ask a question → the model may emit a tool call in a fenced block:

```tool
{"tool": "disk"}
```

Read-only tools run and return results; the model explains them. Mutating
tools (`run_fix`, `run_command`, `install_updates`) prompt `Approve? [y/N]`
first. The model then summarises the result in plain English.

## Modes (config.json)

| mode   | needs                      | notes                                  |
|--------|----------------------------|----------------------------------------|
| local  | a GGUF at `model_path`     | default; private; CPU-only             |
| api    | `OPENAI_API_KEY` / config  | OpenAI-compatible endpoint (any host)  |

Set `"mode": "api"` and your key to use a cloud model instead (e.g. OpenRouter).

## Tools

Read-only (auto): `system_info`, `disk`, `event_log`, `network`, `services`, `run_diagnostic <script>`
Mutating (confirm): `run_fix <script>`, `run_command <powershell>`, `install_updates`

All scripts are discovered from `../Scripts` (the AutoByte library).
