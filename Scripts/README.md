# AutoByte Script Catalog

98 → 93 scripts after cleanup (removed empty `Test.ps1`, typo duplicate
`Get-CurrentLoggedInUuser.ps1`, superseded `M365_Office32bitRepair_copy.ps1`,
redundant `InstallRevit2024-2-0-63.ps1`, and `DEBUGScript.ps1` scratch).

Most scripts are standalone and meant to be run directly (often as Admin). A
few read a path/printer model from `Read-Host` prompts. Grouped by purpose below.

## Printers & Supplies
- `Get-PrinterSupplies.ps1` — read ink/toner levels for networked printers.
- `Get-PrinterTonerLevel.ps1` — toner level report.
- `Get-HPPrinterSupplies.ps1` — HP-specific supply readout.
- `GetUSBPrinter.ps1` — enumerate USB-attached printers.
- `DeployPrinter.ps1` — deploy a printer from a driver ZIP (prompts for path/model).
- `Deploy-Printer.ps1` — alternate printer deployment (prompts for model).
- `InstallPrinter.ps1` — install a printer (third variant).
- `PrinterCleanup.ps1` — remove stale/orphaned printer objects.
- `PrinterFix.ps1` — stop spooler, clear spool files, restart (basic printer fix).

## Office / M365
- `InstallOffice.ps1` — install Office via ODT.
- `InstallMSProjects.ps1` — install MS Project via ODT.
- `M365_Office32bitRepair.ps1` — repair/reinstall 32-bit Office (v1.1, feature-rich).
- `M365_Office64bitRepair.ps1` — 64-bit Office repair.
- `M365-AddTrustedLocation.ps1` / `OfficeSetTrustedLocation.ps1` — set Office trusted locations.
- `Get-ExcelAddIns.ps1` — list Excel add-ins.
- `RemoveAllTeams.ps1` / `InstallTeams.ps1` — Teams removal / install.
- `NewTeamsReadinessCheckScript.ps1` — pre-req check for new Teams.
- `TeamsOutlookTroubleshoot.ps1` — Teams/Outlook integration fixes.

## Windows Repair & Updates
- `CBSRepair.ps1` — DISM/CBS component store repair.
- `WindowsOnlineRepair.ps1` / `WindowsOSRepair.ps1` / `WindowsSystemRepair.ps1` / `Windows_10_Online_Repair_Script.ps1` — online/component repairs (overlapping; keep the one you prefer).
- `InstallWindowsUpdate.ps1` — trigger Windows Update install.
- `RemoveWindowsUpdatePolicy.ps1` — clear update GPO/policy blockers.
- `UpgradeToWin11.ps1` / `UpgradeWindowsToLatestFeature.ps1` — feature upgrades.
- `UpdateWin11To24H2-BypassTpmandCpu.ps1` — 24H2 upgrade with TPM/CPU bypass.
- `ResetandClearWindowsSearchDB.ps1` — rebuild Windows Search index.
- `CheckIfSystemNeedsToRestart.ps1` — pending-reboot check.
- `CheckUserProfileIssue.ps1` — user-profile corruption check.

## Network & Internet
- `InternetSpeedTest.ps1` — bandwidth test.
- `InternetLatencyTest.ps1` / `InternetLatencyTestNew.ps1` — latency probes (two versions).
- `NetworkScan.ps1` — basic network scan.
- `GetWindowsEvents.ps1` / `WinLog.ps1` — pull Windows event logs.
- `SplitlogFile.ps1` — split a large log into chunks.
- `SimpleHttpServer.ps1` — quick local HTTP server (port 8080).

## User / Profile / Admin
- `CheckIfUserIsAdmin.ps1` — admin-elevation check.
- `Get-CurrentLoggedInUser.ps1` (and `GetLoggedInUserName.ps1`) — who is logged on.
- `RenameOnedrive.ps1` — fix OneDrive folder naming.
- `EnableFilesOnDemand.ps1` — enable OneDrive Files On-Demand.
- `CheckIfOneDriveSyncFolder.ps1` — OneDrive sync-folder check.
- `AddFolderToHomeNameSpace.ps1` — add folder to Explorer Home.
- `MigrateToEntraCloud.ps1` — Entra/cloud profile migration helper.

## Dell / HP Hardware
- `DellCommandUpdate.ps1` — run Dell Command Update.
- `DellBloatWareUninstaller.ps1` / `Uninstall-DellApps.ps1` / `BloatWareRemover.ps1` — Dell bloat removal.
- `CheckDellRaid.ps1` — Dell RAID troubleshooting log.
- `CertKeyVerifier.ps1` — cert/key validation (WPF).

## Calendar / Exchange (Exchange Online)
- `Get-CalendarSharingStatus.ps1` / `GetCalendarPermissions.ps1` — read calendar perms.
- `SetCalendarPermissions.ps1` / `SetCalendarToReviewer.ps1` / `SetExchangeCalendarToReViewer.ps1` / `GrantFullAccessToCalendar.ps1` — set calendar perms.
- `MSOutlookCalendarPermissions_Test.ps1` — test/diagnostic for calendar perms.
- `Set-OutlookAsDefaultEmailClient.ps1` — set Outlook as default mail client.
- `Set-SmtpRelay.ps1` — SMTP relay config helper.
- `3CXProv.ps1` — 3CX provisioning.
- `GSuite_WorkSpace_Diag.ps1` — Google Workspace diagnostic.

## Deployment / Install
- `DownloadFile.ps1` / `DownloadandInstallPackage.ps1` — download + install helpers.
- `InstallMSIEXE.ps1` — generic MSI/EXE install.
- `Install-Winget.ps1` / `InstallWinGet.ps1` — Winget bootstrap (two versions).
- `Fido.ps1` — Windows ISO downloader (Fido).
- `InstallRevit2024.2.0.63.ps1` — Revit 2024 install.
- `Add-CertChain.ps1` — install a certificate chain.
- `FinancialApp.ps1` / `TestSage.ps1` / `QBDiag.ps1` / `WorkPaperMonitorTroubleShooter.ps1` / `CCHUninstall.ps1` — LOB app diagnostics (Sage, QuickBooks, CCH, Workpaper Monitor).
- `BuildingTestApp.ps1` — app build/test helper.

## Diagnostics / Utilities
- `MemoryAccessViolation.ps1` — memory-access-violation troubleshooting.
- `IntuneRemediate.ps1` — Intune remediation helper.
- `LanguageIDs.ps1` / `LanguagePacks.ps1` — language pack helpers.
- `DiskCleaner.ps1` / `CheckDriveSpace.ps1` — disk cleanup / space check.
- `DebugScript.ps1` — REMOVED (was hello-world scratch).
- `TestAI.ps1` — PSLLM local-model experiment (keep).
- `Testing.ps1` / `ScriptStar.ps1` — alternate launchers / scratch (kept).

## macOS (shell)
- `[MAC]DIskCleanup.sh`, `install_anydesk_macos.sh`, `install_screenconnect_macos.sh` — macOS helpers.

## Windows batch
- `Skip_TPM_Check_on_Dynamic_Update.cmd` — skip TPM check on dynamic update.

## Launchers
- `AutoByte.ps1` (repo root) — original hub: enforces TLS1.2, lists a curated
  set of scripts by URL, runs them.
- `ScriptStar.ps1` — alternate hub with a different curated list.
- `AutoByteGUI.py` + `AutoByteGUI-Build.ps1` — Python/Tk GUI front-end.

> NOTE: scripts are run-at-your-own-risk MSP utilities. Many require Admin and
> touch system settings. Review before production use.
