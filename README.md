# AutoByte

**Windows troubleshooting and repair toolkit for IT technicians and MSPs.**
A curated library of PowerShell repair and diagnostic scripts, a safe launcher that runs them on any PC
or through your RMM, and a local AI assistant that diagnoses a PC in plain English.

![AutoByte](AutoByte_Main.png)

## New: AutoByte v3 (one engine, one command)

```powershell
abyte              # plain-language PC check-up for anyone
abyte scan --report --open
abyte fix --suggested --whatif
abyte sandbox      # try fixes in a throw-away copy of Windows first
abyte remote       # let IT support connect (your support page or Quick Assist)
```

Read-only health checks (Event Viewer analysis, "why is it slow?", drivers, missing patches and
CVE exposure, outdated apps), safe self-verifying fixes, an HTML report, JSON for Datto RMM, and the
optional AI agent on top. Everything below (v2) keeps working. Install, Datto setup and how it
works: **[docs/V3.md](docs/V3.md)**.

## What's in it

| Tool | What it does | Who it's for |
|---|---|---|
| **Script launcher** (`AutoByte.ps1`) | Menu of common fixes (disk cleanup, Windows Update reset, system repair, Office repair, network tests…). Every script is downloaded from a pinned release and SHA-256 verified before it runs. | Technicians, RMM jobs |
| **Script library** (`Scripts/`) | Windows repair, Office/M365, Exchange calendar permissions, printers, Dell/HP, Teams, network, deployment. Listed with descriptions in [`scripts.manifest.json`](scripts.manifest.json). | Anyone |
| **CLI** (`AutoByte-CLI.ps1`) | Browse, search and run any script from a local copy. | Technicians |
| **AI agent** (`AutoByte-Agent/`) | Describe a problem in plain English; it checks the PC, explains what's wrong and proposes a fix. Runs a small model **on your PC** (no data leaves it, no API bill), or a cloud model if you prefer. | Technicians, power users |
| **GUI** (`AutoByteGUI.py`) | Buttons for the most-used scripts, with a live output window. | Anyone who prefers clicking |

## Install and run

Requirements: Windows 10 or 11 and an **administrator** PowerShell window (right-click PowerShell →
*Run as administrator*). Windows PowerShell 5.1 is built in; nothing else is needed for the launcher.

### Script launcher (fastest way to start)

```powershell
$f = Join-Path $env:TEMP 'AutoByte.ps1'
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/deep1ne8/misc/v2.0.0/AutoByte.ps1' -OutFile $f -UseBasicParsing
powershell -NoProfile -ExecutionPolicy Bypass -File $f
```

Pick a number from the menu. To run one script without the menu (for RMM), add `-Script <Name>`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File $f -Script CheckDriveSpace
```

The launcher exits with the script's exit code: `0` success, `1` failure, `3010` reboot required.
Add `-WhatIf` to download and verify without running.

### AI agent

Needs [Python 3.11+](https://www.python.org/downloads/) (tick *Add python.exe to PATH* during setup).

```powershell
git clone https://github.com/deep1ne8/misc.git
cd misc\AutoByte-Agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python agent.py --mock          # try it instantly, no download
```

To use the real local model (~4.7 GB download, 16 GB RAM recommended):

```powershell
pip install -r requirements.txt -r requirements-setup.txt
python setup_model.py           # downloads and verifies the model
python agent.py
```

Then type a problem, for example *"my PC is slow"* or *"is a restart pending?"*. Read-only checks run on
their own; anything that changes the PC shows exactly what it will do and waits for you to type `y`.

### GUI

```powershell
git clone https://github.com/deep1ne8/misc.git
cd misc
powershell -ExecutionPolicy Bypass -File .\AutoByteGUI-Build.ps1
```

This installs Python if needed and opens the window. Next time, just run `python AutoByteGUI.py`.

### CLI

```powershell
git clone https://github.com/deep1ne8/misc.git
cd misc
powershell -ExecutionPolicy Bypass -File .\AutoByte-CLI.ps1        # numbered menu
powershell -ExecutionPolicy Bypass -File .\AutoByte-CLI.ps1 list   # every script with a description
```

**Full guide** (RMM variables, cloud-model mode, every option):
[docs/GETTING-STARTED.md](docs/GETTING-STARTED.md)

## Safety

- **Verified downloads.** The launcher and GUI only run scripts from a pinned release whose SHA-256
  matches `scripts.manifest.json`. No `Invoke-Expression` of downloaded code.
- **Nothing changes without your OK.** The AI agent runs only read-only checks by itself. Repairs need
  an explicit `y`, and destructive commands (disk format, boot changes, wiping system folders…) are
  refused outright. Every action is logged to `%ProgramData%\AutoByte\Logs`.
- **No secrets in the repo.** Scripts that need keys (Liongard, ImmyBot) read them from parameters or
  environment variables, such as your RMM's secure variables.
- **Unsupported scripts are fenced off.** Scripts that bypass Windows 11 requirements are flagged and
  need `-AllowUnsupported`; the agent can't run them.

## Project layout

```
AutoByte.ps1            launcher (menu + RMM mode)
AutoByte-CLI.ps1        local script browser
AutoByteGUI.py          GUI  (AutoByteGUI-Build.ps1 installs it)
AutoByte-Agent/         AI troubleshooting agent
Scripts/                script library (see scripts.manifest.json)
archive/                retired scripts, not runnable by the tools
src/AutoByte/           v3 engine: module, tasks, adapters
bin/abyte.ps1           v3 command line (end-user menu, scan, fix, report)
install/                v3 hash-verified installer
datto/                  v3 Datto RMM component
build/                  manifest generator, local CI, leaked-key guard
docs/                   getting-started guide, changelog
tests/                  Pester and pytest suites
```

## Contributing

Run the same checks as CI on your own PC (free, no cloud account needed):

```powershell
pwsh ./build/Invoke-LocalCI.ps1 -InstallHooks   # once: checks run automatically on commit and push
pwsh ./build/Invoke-LocalCI.ps1                 # run everything now (shows progress for each step)
```

After changing anything in `Scripts/`, regenerate the manifest: `pwsh ./build/Build-ScriptManifest.ps1`.
Use [Conventional Commits](https://www.conventionalcommits.org/) (`fix(scripts): …`, `feat(agent): …`).

## Support

- Bugs and requests: [GitHub Issues](https://github.com/deep1ne8/misc/issues)
- Release notes: [docs/changelog.md](docs/changelog.md)
