# AutoByte — MSP Automation Toolkit

> **Production-grade PowerShell automation for MSP engineers.**  
> Covers system health, Windows Update, disk cleanup, software deployment, M365 diagnostics, Intune sync, and more — all from a single interactive menu.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-0078d4?style=flat-square&logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11%20%7C%20Server-555?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Version](https://img.shields.io/badge/Version-2.0.0-00d4ff?style=flat-square)

---

## What it does

AutoByte is a self-contained PowerShell toolkit that replaces a pile of one-off scripts with a single, structured tool any MSP engineer can run on a client machine. No dependencies to pre-install. No GUI frameworks. Just PowerShell.

| Module | What it automates |
|--------|-------------------|
| **System Health** | CPU, RAM, disk usage, uptime — at a glance |
| **Windows Update** | Scan and install via PSWindowsUpdate |
| **Disk Cleanup** | Temp files, WU cache, IE cache — shows MB freed |
| **Software Deploy** | 10 common apps via `winget` — silent, no prompts |
| **M365 Diagnostics** | Module presence check + connectivity test |
| **OneDrive FOD** | Enable/disable Files On-Demand via registry |
| **Dell Command Update** | Trigger `dcu-cli.exe /applyUpdates` |
| **Network Diagnostics** | Ping key endpoints, list active adapters |
| **Intune Sync** | Trigger MDM + Intune Management Extension sync |
| **HTML Report** | Full system snapshot exported to Desktop |

---

## Quick Start

```powershell
# Clone
git clone https://github.com/deep1ne8/misc.git
cd misc

# Run (requires Administrator)
Set-ExecutionPolicy -Scope Process Bypass
.\AutoByte.ps1
```

> **Requires:** PowerShell 5.1+, Windows 10/11 or Server 2016+, run as Administrator.

---

## Usage

AutoByte presents an interactive numbered menu. Select a module by number and follow any prompts. All actions are logged to:

```
C:\ProgramData\AutoByte\Logs\AutoByte_YYYYMMDD.log
```

### HTML Report

Option `10` generates a self-contained HTML report saved to your Desktop — useful for client documentation or ticket attachments.

---

## Modules

### System Health
Pulls CPU load percentage, RAM usage (used/total GB), system uptime, and per-drive disk usage — all in one shot.

### Windows Update
Installs the `PSWindowsUpdate` module if not present, then runs a full scan and install pass. Reboots are suppressed by default.

### Disk Cleanup
Targets `%TEMP%`, `%SystemRoot%\Temp`, Windows Update download cache, and IE cache. Reports total MB freed.

### Software Deployment
Wraps `winget install` for 10 common MSP software titles. Silent install, no user interaction required. Extend `$apps` in the script to add your own.

Available titles:
- 7-Zip, Chrome, Firefox, Notepad++, VLC
- VS Code, Zoom, Microsoft Teams, Adobe Reader, Greenshot

### M365 Diagnostics
Checks whether key PowerShell modules are installed (`MSOnline`, `AzureAD`, `ExchangeOnlineManagement`, `MicrosoftTeams`) and tests connectivity to `login.microsoftonline.com`.

### OneDrive Files On-Demand
Writes `FilesOnDemandEnabled` DWORD to `HKCU:\Software\Microsoft\OneDrive` — no OneDrive restart required.

### Dell Command Update
Calls `dcu-cli.exe /applyUpdates -reboot=disable`. Fails gracefully if Dell Command Update is not installed.

### Network Diagnostics
Pings `8.8.8.8`, `1.1.1.1`, `login.microsoftonline.com`, and `outlook.office365.com` with latency. Lists all active network adapters and their link speeds.

### Intune Sync
Triggers the Intune Management Extension sync agent and the `PushLaunch` scheduled task. Works on Intune-enrolled devices.

---

## Extending AutoByte

Add a new module in three steps:

1. Write a `function Invoke-YourModule` block using the `Invoke-WithLog` wrapper.
2. Add a label string to `$menuItems` in `Start-AutoByte`.
3. Add a `case` entry in the `switch` block.

```powershell
function Invoke-YourModule {
    Invoke-WithLog 'Your module description' {
        # your code here
        Write-Log 'Done.' -Level OK
    }
}
```

---

## Logging

Every action writes to a date-stamped log at `C:\ProgramData\AutoByte\Logs\`. Log levels: `INFO`, `WARN`, `ERROR`, `OK`.

---

## Requirements

| Requirement | Minimum |
|-------------|---------|
| PowerShell | 5.1 |
| OS | Windows 10 / Windows Server 2016 |
| Privileges | Administrator |
| Winget | Required for Software Deployment module only |

---

## Contributing

Pull requests welcome. For major changes, open an issue first.

1. Fork the repo
2. Create a feature branch: `git checkout -b feature/my-module`
3. Commit: `git commit -m 'Add: my-module'`
4. Push and open a PR

---

## License

MIT — free to use, modify, and distribute. Attribution appreciated.

---

*Built by an MSP engineer, for MSP engineers.*  
[github.com/deep1ne8/misc](https://github.com/deep1ne8/misc)
