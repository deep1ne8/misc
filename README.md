# AutoByte Agent

A **lightweight, local Windows AI troubleshooting agent**. Talk to it in plain
English and it will diagnose, explain, repair, and update your Windows PC —
and answer any Windows question. It runs a small AI model **entirely on your
machine** (CPU-only, no GPU needed): your data never leaves the PC, and there
is **no token limit** and no API bill.

![AutoByte Agent](https://github.com/deep1ne8/misc/blob/main/AutoByte_Main.png)

## Why

Most "AI PC assistants" send everything to the cloud. AutoByte Agent doesn't.
It loads a small open-source model locally and drives a **safe tool layer**
that inspects your system with PowerShell probes and — only with your explicit
approval — runs fixes from the bundled script library.

## How it works

1. You describe a problem or ask a question in plain English.
2. The agent reasons and may call a tool (e.g. check disk, read event logs,
   inspect a service). **Read-only diagnostics run automatically.**
3. If it thinks a change is needed (repair, update, command), it shows exactly
   what it will do and asks **`Approve? [y/N]`**. Nothing mutates your system
   without your `Y`.
4. It explains the result in plain English.

## Quick start

```powershell
git clone https://github.com/deep1ne8/misc.git
cd misc\AutoByte-Agent
pip install -r requirements.txt

# one-time model download (~4.7GB, runs on CPU):
python setup_model.py

# chat with your PC:
python agent.py
```

No model yet? Test the whole framework (tool-calling + safety gate) with no
download:

```powershell
python agent.py --mock
```

## Modes (`AutoByte-Agent/config.json`)

| mode  | needs                       | notes                                       |
|-------|-----------------------------|---------------------------------------------|
| local | a GGUF at `model_path`      | default — private, CPU-only, no token limit |
| api   | `OPENAI_API_KEY` / config   | OpenAI-compatible endpoint (any host)       |

Set `"mode": "api"` and your key to use a cloud model instead.

## Tools

**Read-only (auto-run):** `system_info`, `disk`, `event_log`, `network`,
`services`, `run_diagnostic <script>`

**Mutating (confirm first):** `run_fix <script>`, `run_command <powershell>`,
`install_updates`

The agent can call any of the **93 bundled PowerShell scripts** in `../Scripts`
(printers, Office/M365, Windows repair, network, Dell/HP, calendar/Exchange,
deploy). Full detail in [`AutoByte-Agent/README.md`](AutoByte-Agent/README.md).

## Safety

- Mutating actions never run without an explicit `Y`.
- No tool formats disks or bulk-deletes data; repairs are bounded scripts or
  commands you approve line-by-line.
- Everything is local; no telemetry, no cloud calls (unless you opt into api mode).

## Project layout

```
misc/
  AutoByte-Agent/   <- the AI agent (this is the project)
    agent.py        <- orchestrator + tool-calling loop
    tools.py        <- safe Windows tool layer
    setup_model.py  <- one-command local model download
    config.json     <- model mode / path / api key
    README.md       <- agent docs
  Scripts/          <- 93 bundled PowerShell scripts the agent can call
  README.legacy.md  <- archived pre-agent README
```

## Support

- Issues: [GitHub Issues](https://github.com/deep1ne8/misc/issues)
- Contributions welcome. 🎉
