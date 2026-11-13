# 📜 AutoByte Release Notes

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