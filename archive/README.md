# Archive

Scripts kept for reference only. They are **not** in `scripts.manifest.json`, so the
launcher, GUI, CLI and agent cannot run them. Do not deploy them via RMM.

| File | Why archived |
|---|---|
| `ScriptStar.ps1` | Legacy launcher: runs scripts with `Invoke-Expression` from `main`, several URLs 404. Replaced by `AutoByte.ps1`. |
| `TestAI.ps1` | PSLLM experiment; superseded by `AutoByte-Agent/`. |
| `FinancialApp.ps1` | Deploys an unrelated web app; not an MSP endpoint script. |
| `MSOutlookCalendarPermissions_Test.ps1` | Test copy of the calendar-permission scripts (`SetCalendarPermissions.ps1` is canonical). |
