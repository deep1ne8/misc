# 📜 AutoByte Release Notes

## 🎨 Version 3.1.0 (unreleased)
- Redesigned console output: pastel colour palette (24-bit where supported, 16-colour and plain
  fallbacks), plain-language statuses, grouped findings, wrapped details and clear next steps.
- New end-user menu with options to try fixes in **Windows Sandbox** and to **get remote help**
  (MSP support page via `RemoteSupportUrl`, else Quick Assist; mentions Datto Web Remote).
- `abyte sandbox [command]` / `--enable` / `--no-network`; `abyte remote`; `--plain`, `--ascii`.
- `update-apps`: progress per app, plain reasons for winget errors (app open, per-user install,
  broken uninstaller, ...), per-app results in `Data`, and a new **Partial** status (exit code 1)
  when some apps update and others do not. Fixed "winget not available" being reported after
  every app had been updated.
- Reports honour `AUTOBYTE_REPORT_DIR`.

## 🧩 Version 3.0.0
- New engine: `src/AutoByte` module with 19 read-only checks and 14 idempotent, self-verifying fixes.
- Event Viewer analysis across System, Application, Setup, Defender and boot-performance logs, with
  blue-screen stop codes and Windows Update error codes mapped to the right fix.
- "Why is it slow?" performance check, driver problems, missing security updates, CVE exposure
  (MSRC fixed builds + CISA KEV, refreshed daily) and outdated apps (winget), each with a fix.
- `abyte` CLI: end-user menu, `scan`, `fix`, `list`, `report`, `legacy`; JSON output; exit codes 0/1/2/3010.
- Self-contained HTML health report.
- Hash-verified installer (`install/Install-AutoByte.ps1`) and Datto RMM component (UDF + monitor output).
- AI agent: `health_scan` and `run_task` tools on top of the engine.
- v2 launcher and scripts unchanged; `abyte legacy <name>` bridges to them.

## 🔐 Version 2.0.0 (unreleased)
**Security**
- Removed embedded Liongard and ImmyBot credentials; values are now parameters/env vars.
- `AutoByte.ps1` and the GUI download from a pinned release, verify SHA-256 against
  `scripts.manifest.json`, and never use `Invoke-Expression`.
- Machine-wide TLS registry change is opt-in (`-PersistStrongCrypto`).
- AutoByte Agent: `run_diagnostic` can only auto-run manifest-approved read-only scripts;
  destructive commands are refused; every mutation is audit-logged.

**Fixes**
- Agent no longer crashes on plain-English answers.
- Launcher menu item 14 (Office 64-bit repair) is reachable again.
- `QBDiag.ps1` and `TeamsOutlookTroubleshoot.ps1` run on Windows PowerShell 5.1 (no ternaries).
- `ExchangeMailboxHealthCheck.ps1` parses (13 syntax errors fixed).

**Repo / CI**
- Generated `scripts.manifest.json`; vendor binaries removed from git.
- CI: PSScriptAnalyzer (+5.1 parse), Pester, ruff/mypy/pytest on Windows and Linux,
  gitleaks, CodeQL, Dependabot. All actions pinned to commit SHAs.

## 🚀 Version 1.0.1
**🔹 Features:**
- Added Chocolatey package support
- Improved installation and execution

## ✅ Version 1.0.0
- Initial release with core automation features