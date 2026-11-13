# AutoByte Agent

A **lightweight, local Windows AI troubleshooting agent**. It talks to you in
plain English, inspects your PC with safe PowerShell probes, explains what's
wrong, and — with your explicit approval — runs fixes from the bundled
AutoByte script library or its own repair commands.

- **Local & private**: runs a small GGUF model on your CPU (no data leaves the machine, no token limit).
- **Safe by design**: read-only diagnostics run automatically; anything that *changes* your system (repairs, updates, arbitrary commands) requires you to type `Y`.
- **Tool-using**: reasons over `system_info`, `disk`, `event_log`, `network`, `services`, and the scripts listed in `../scripts.manifest.json`.

## Quick start

```powershell
cd AutoByte-Agent
pip install -r requirements.txt            # local mode
pip install -r requirements-setup.txt      # for setup_model.py
# pip install -r requirements-api.txt      # only for api mode

# 1) get a local model (~4.7GB, one-time, SHA-256 verified):
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

`config.json` is created from `config.example.json` and is git-ignored. API keys
are read from the `OPENAI_API_KEY` environment variable, never from the file.

| mode   | needs                      | notes                                  |
|--------|----------------------------|----------------------------------------|
| local  | a GGUF at `model_path`     | default; private; CPU-only             |
| api    | `OPENAI_API_KEY` / config  | OpenAI-compatible endpoint (any host)  |

Set `"mode": "api"` and your key to use a cloud model instead (e.g. OpenRouter).

## Tools

Read-only (auto): `system_info`, `disk`, `event_log`, `network`, `services`, `run_diagnostic <script>`
Mutating (confirm): `run_fix <script>`, `run_command <powershell>`, `install_updates`

## Safety gate

- Script names must match a manifest entry **exactly**; every script is
  SHA-256-verified against `scripts.manifest.json` before it runs.
- `run_diagnostic` only runs scripts the manifest marks `agentDiagnostic`
  (non-mutating, non-interactive). Anything else must go through `run_fix` (approval).
- Scripts flagged `unsupported` (e.g. Windows 11 TPM/CPU bypass) are never runnable by the agent.
- `run_command` refuses destructive patterns outright (disk format/clear, bcdedit,
  `cipher /w`, recursive deletes of system/user/registry roots, `iex`/download-and-run,
  encoded commands, shadow-copy deletion, security-tool disablement, forced reboot).
  The deny-list is defence in depth; operator approval remains the primary gate.
- Every proposal, refusal, decline and executed mutation is logged as JSON lines to
  `%ProgramData%\AutoByte\Logs\Agent_<yyyyMMdd>.log` (override with `AUTOBYTE_LOG_DIR`).

## Tests

```powershell
pip install -r requirements-dev.txt
python -m pytest; ruff check .; mypy
```
