# Getting started with AutoByte

AutoByte has four parts. Pick the one you need.

| You want to… | Use | Section |
|---|---|---|
| Run a repair/diagnostic script on a PC (by hand or from RMM) | `AutoByte.ps1` launcher | [1](#1-launcher-menu-and-rmm) |
| Browse and run any script from a local copy | `AutoByte-CLI.ps1` | [2](#2-cli-from-a-local-copy) |
| Ask an AI assistant to diagnose a PC in plain English | `AutoByte-Agent` | [3](#3-ai-troubleshooting-agent) |
| Click buttons instead of typing | `AutoByteGUI.py` | [4](#4-gui) |
| Change code in this repo | local CI + git hooks | [5](#5-developers-local-ci-no-cloud-billing) |

Requirements: Windows 10/11, Windows PowerShell 5.1 (built in) or PowerShell 7.4+. Most scripts need
an **elevated** (Run as administrator) prompt.

---

## 1. Launcher (menu and RMM)

Downloads scripts from the pinned release `v2.0.0`, checks each one's SHA-256 against
`scripts.manifest.json`, and only then runs it.

**Interactive menu** (elevated PowerShell):

```powershell
$f = Join-Path $env:TEMP 'AutoByte.ps1'
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/deep1ne8/misc/v2.0.0/AutoByte.ps1' -OutFile $f -UseBasicParsing
powershell -NoProfile -ExecutionPolicy Bypass -File $f
```

Pick a number; `15` exits.

**One script, no menu (RMM / scheduled task).** The launcher exits with the script's exit code
(0 success, 1 failure, 3010 reboot required):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File $f -Script CheckDriveSpace
```

Useful switches:

| Switch | Effect |
|---|---|
| `-WhatIf` | Download and verify only; do not run |
| `-Ref <tag or 40-char SHA>` | Use another release (default `v2.0.0`) |
| `-PersistStrongCrypto` | Also enable .NET strong crypto machine-wide (HKLM). Off by default |
| `-AllowUnsupported` | Allow scripts flagged unsupported (e.g. Windows 11 TPM bypass) |

Script names are the file names without `.ps1`; see `scripts.manifest.json` for the full list.

**Scripts that need credentials** read them from parameters or environment variables, never from the
file. Set these as secure/masked variables in your RMM:

| Script | Variables |
|---|---|
| `agent-installer.ps1` (Liongard) | `LIONGARD_URL`, `LIONGARD_ACCESS_KEY_ID`, `LIONGARD_ACCESS_KEY_SECRET`, `LIONGARD_ENVIRONMENT` |
| `IntuneRemediate.ps1` (ImmyBot) | `IMMY_BASE_URL`, `IMMY_INSTALL_ID`, `IMMY_INSTALL_KEY` |
| `Set-OutlookAsDefaultEmailClient.ps1` | parameter `-TargetEmailAddress` |

## 2. CLI from a local copy

```powershell
git clone https://github.com/deep1ne8/misc.git
cd misc
.\AutoByte-CLI.ps1              # numbered menu of every script
.\AutoByte-CLI.ps1 list         # names + descriptions
.\AutoByte-CLI.ps1 categories   # grouped
.\AutoByte-CLI.ps1 run CheckDriveSpace
```

If PowerShell blocks the file, run once: `Unblock-File .\AutoByte-CLI.ps1` or start it with
`powershell -ExecutionPolicy Bypass -File .\AutoByte-CLI.ps1`.

## 3. AI troubleshooting agent

Needs **Python 3.11+** (python.org; tick "Add python.exe to PATH").

```powershell
cd misc\AutoByte-Agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Try it with no model (30 seconds):**

```powershell
python agent.py --mock
```

Type `check my disk`, `is a restart pending`, `my pc is slow` (answer `n`), `wipe it` (refused), `quit`.

**Run it with a local model** (private, CPU-only, no API bill; ~4.7 GB download, 16 GB RAM recommended):

```powershell
pip install -r requirements.txt -r requirements-setup.txt
python setup_model.py           # downloads + SHA-256 verifies the model, writes config.json
python agent.py
```

`llama-cpp-python` compiles on install if no prebuilt wheel matches your Python. If that fails, install
"Visual Studio Build Tools" (C++ workload) and retry, or use API mode.

**Or use a cloud model (API mode):**

```powershell
pip install -r requirements-api.txt
copy config.example.json config.json   # then set "mode": "api" (and "model", "api_base" if not OpenAI)
$env:OPENAI_API_KEY = '<your key>'     # never put the key in config.json
python agent.py
```

Read-only checks run automatically. Anything that changes the PC shows exactly what it will do and waits
for `y`. Destructive commands are refused outright. Every proposal and action is logged to
`%ProgramData%\AutoByte\Logs\Agent_<date>.log`.

## 4. GUI

```powershell
cd misc
powershell -ExecutionPolicy Bypass -File .\AutoByteGUI-Build.ps1
```

Installs Python if missing (signature-checked) and the pinned `requests`, then opens the window. Next
time, just run `python AutoByteGUI.py`. Like the launcher, it only runs hash-verified scripts from `v2.0.0`.

## 5. Developers: local CI (no cloud billing)

Everything the GitHub pipeline checks also runs on your own PC, free:

```powershell
cd misc
pwsh ./build/Invoke-LocalCI.ps1 -InstallHooks   # once per clone
pwsh ./build/Invoke-LocalCI.ps1                 # leaked-key guard, manifest, parse (7.x + 5.1), PSScriptAnalyzer, Pester, ruff, mypy, pytest
```

After `-InstallHooks`:

- **pre-commit** refuses a commit that contains any previously leaked credential
  (`build/Test-LeakedSecrets.ps1`, matching SHA-256 fingerprints in `build/leaked-secrets.sha256`).
- **pre-push** runs the full local CI and refuses the push if anything fails.

Skip a slow step while iterating: `pwsh ./build/Invoke-LocalCI.ps1 -Skip Python,Analyzer`.
After editing anything in `Scripts/`, regenerate the manifest: `pwsh ./build/Build-ScriptManifest.ps1`.

GitHub Actions is optional. This repository is public, so standard runners are free. Every workflow job
also skips itself if the repository is ever made private, so it can never use paid minutes.

**Releasing:**
1. In a PR, set the new version as the default: `-Ref` in `AutoByte.ps1`, `REF` in `AutoByteGUI.py` and
   `$script:Ref` in `tests/powershell/Repository.Tests.ps1`. Run local CI, then merge.
2. Tag the merged commit and push the tag: `git tag -a v2.0.1 -m "AutoByte v2.0.1"` then `git push origin v2.0.1`.
   Endpoints switch to the new release only once that tag exists.
