# Event log knowledge base: which events matter, what they mean, and which fix task addresses them.
# Each rule: Code, Log, Provider (regex on ProviderName), Id, Severity, Title, Detail, FixTask,
# optional MinCount (default 1), Level (default 1,2,3 = Critical/Error/Warning; Information events
# are never problems, e.g. Ntfs 98 "volume is healthy"), Group (scriptblock -> group key, one finding
# per group), Fix (scriptblock -> fix task id, overrides FixTask per group) and Summarize
# (scriptblock over the group's events -> text appended to the detail).
# Keep rules locale-independent: match provider + id, read named EventData fields, never message text.

$script:AbStopCodes = @{
    '0x0000000A' = @('IRQL_NOT_LESS_OR_EQUAL', 'A driver accessed memory it should not.', 'install-driver-updates')
    '0x0000001A' = @('MEMORY_MANAGEMENT', 'Memory corruption. Test the RAM (mdsched.exe) and update drivers.', 'install-driver-updates')
    '0x0000003B' = @('SYSTEM_SERVICE_EXCEPTION', 'Usually a faulty driver or a damaged system file.', 'install-driver-updates')
    '0x00000050' = @('PAGE_FAULT_IN_NONPAGED_AREA', 'A driver or faulty RAM referenced invalid memory.', 'install-driver-updates')
    '0x0000007A' = @('KERNEL_DATA_INPAGE_ERROR', 'Windows could not read from disk. Check the disk and cabling.', 'repair-disk-errors')
    '0x0000007E' = @('SYSTEM_THREAD_EXCEPTION_NOT_HANDLED', 'A driver crashed in a system thread.', 'install-driver-updates')
    '0x0000009F' = @('DRIVER_POWER_STATE_FAILURE', 'A driver failed during sleep/wake. Update drivers and BIOS.', 'install-driver-updates')
    '0x000000D1' = @('DRIVER_IRQL_NOT_LESS_OR_EQUAL', 'A driver (often network or storage) used bad memory.', 'install-driver-updates')
    '0x000000EF' = @('CRITICAL_PROCESS_DIED', 'A core Windows process stopped. System files may be damaged.', 'repair-system-files')
    '0x00000116' = @('VIDEO_TDR_FAILURE', 'The graphics driver stopped responding.', 'install-driver-updates')
    '0x00000124' = @('WHEA_UNCORRECTABLE_ERROR', 'A hardware fault (CPU, RAM, board). Update BIOS; check cooling.', 'install-driver-updates')
    '0x00000133' = @('DPC_WATCHDOG_VIOLATION', 'A driver (often storage/SSD firmware) took too long.', 'install-driver-updates')
    '0x00000139' = @('KERNEL_SECURITY_CHECK_FAILURE', 'A driver or system file corrupted kernel data.', 'repair-system-files')
    '0x0000001E' = @('KMODE_EXCEPTION_NOT_HANDLED', 'A kernel-mode driver crashed.', 'install-driver-updates')
}

# Windows Update error -> most likely fix.
$script:AbUpdateErrors = @{
    '0x800F081F' = @('Component store is damaged (source files missing).', 'repair-system-files')
    '0x800F0831' = @('Component store is missing a previous update.', 'repair-system-files')
    '0x80073712' = @('Component store is damaged.', 'repair-system-files')
    '0x800F0922' = @('Servicing failed (often the System Reserved partition or VPN).', 'repair-system-files')
    '0x80070002' = @('Update files are missing or damaged.', 'reset-windows-update')
    '0x80070070' = @('Not enough disk space.', 'clean-temp-files')
    '0x8024402C' = @('Cannot reach the update server (proxy or DNS).', 'reset-network-cache')
    '0x80072EFD' = @('Cannot connect to the update server.', 'reset-network-cache')
    '0x80072EE2' = @('Connection to the update server timed out.', 'reset-network-cache')
    '0x80240034' = @('The download failed.', 'reset-windows-update')
    '0x8024A105' = @('Windows Update agent error.', 'reset-windows-update')
    '0x80070643' = @('Installation failed (for KB5034441 this means the recovery partition is too small).', '')
    '0x80073D02' = @('A Microsoft Store app could not update because it was open. It retries automatically; close the app or restart.', '')
    '0x80073CF9' = @('A Microsoft Store app failed to install. It usually retries; reset the Store (wsreset) if it persists.', '')
    '0x80073CF3' = @('A Microsoft Store app update failed a dependency check. It usually resolves on a later attempt.', '')
    '0x80073D06' = @('A newer version of the Store app is already installed. Harmless.', '')
}

# Services the start-core-services fix manages (display names as logged by Service Control Manager).
$script:AbCoreServiceNames = '(?i)^(DNS Client|DHCP Client|Windows Event Log|Workstation|Windows Defender Firewall|Windows Management Instrumentation|Task Scheduler|Cryptographic Services|Windows Audio|Dnscache|Dhcp|EventLog|LanmanWorkstation|mpssvc|Winmgmt|Schedule|CryptSvc|AudioSrv)$'

# Windows components: when these crash, system files are the likely cause, not the app.
$script:AbWindowsApps = @('explorer.exe', 'svchost.exe', 'dwm.exe', 'ShellExperienceHost.exe', 'StartMenuExperienceHost.exe',
    'SearchHost.exe', 'SearchApp.exe', 'RuntimeBroker.exe', 'sihost.exe', 'ctfmon.exe', 'taskhostw.exe', 'spoolsv.exe', 'lsass.exe', 'winlogon.exe')

$script:AbEventRules = @(
    # ---- crashes and power ---------------------------------------------------------------------
    @{ Code = 'events.bugcheck'; Log = 'System'; Provider = 'BugCheck|WER-SystemErrorReporting'; Id = 1001; Severity = 'Critical'
        Title = 'Blue-screen crash'; Detail = 'Windows stopped with a STOP error.'
        Group = { param($e) if ((($e.Props -join ' ') + ' ' + ($e.Data.Values -join ' ')) -match '0x[0-9a-fA-F]{8}') { '0x' + $Matches[0].Substring(2).ToUpperInvariant() } else { 'unknown' } }
        Fix = { param($g) if ($script:AbStopCodes.ContainsKey($g)) { $script:AbStopCodes[$g][2] } else { 'install-driver-updates' } }
        Explain = { param($g) if ($script:AbStopCodes.ContainsKey($g)) { "$($script:AbStopCodes[$g][0]): $($script:AbStopCodes[$g][1])" } else { "Stop code $g. Update drivers and BIOS first." } } }
    @{ Code = 'events.unexpected-shutdown'; Log = 'System'; Provider = 'Kernel-Power|EventLog'; Id = 41, 6008; Severity = 'Warning'
        Title = 'Unexpected shutdown'; Detail = 'Power loss, a forced power-off, or a crash without a dump. Check power, battery and recent driver changes.' }

    # ---- storage --------------------------------------------------------------------------------
    @{ Code = 'storage.bad-blocks'; Log = 'System'; Provider = '^disk$|Microsoft-Windows-Disk'; Id = 7, 51, 153; Severity = 'Critical'
        Title = 'Disk read/write errors'; Detail = 'The disk reported bad blocks or retried I/O. Back up now; the disk may be failing.'; FixTask = 'repair-disk-errors' }
    @{ Code = 'storage.filesystem-corrupt'; Log = 'System'; Provider = 'Ntfs'; Id = 55, 98, 130; Level = 1, 2; Severity = 'Critical'
        Title = 'File system corruption'; Detail = 'NTFS found damaged structures on a volume.'; FixTask = 'repair-disk-errors' }
    @{ Code = 'storage.controller-reset'; Log = 'System'; Provider = 'stor(ahci|nvme|port)|iaStor|nvme'; Id = 129, 153; Severity = 'Warning'; MinCount = 3
        Title = 'Storage controller resets'; Detail = 'The disk stopped responding and was reset. Causes freezes. Update storage drivers and SSD firmware.'; FixTask = 'install-driver-updates' }

    # ---- hardware and drivers ----------------------------------------------------------------
    @{ Code = 'hardware.whea'; Log = 'System'; Provider = 'WHEA-Logger'; Id = 1, 17, 18, 19, 47; Severity = 'Critical'
        Title = 'Hardware error (WHEA)'; Detail = 'The CPU, memory or PCIe bus reported an error. Update BIOS and chipset drivers; check cooling and RAM.'; FixTask = 'install-driver-updates' }
    @{ Code = 'drivers.display-reset'; Log = 'System'; Provider = '^Display$'; Id = 4101; Severity = 'Warning'
        Title = 'Graphics driver stopped responding'; Detail = 'The display driver was reset (screen flicker or freeze).'; FixTask = 'install-driver-updates' }
    @{ Code = 'drivers.load-failed'; Log = 'System'; Provider = 'Kernel-PnP'; Id = 219; Severity = 'Info'; MinCount = 2
        Title = 'Drivers loaded late at start-up'; Detail = 'Usually harmless when the devices work; the drivers check lists devices that are actually failing.'
        Summarize = { param($events) $d = @($events | ForEach-Object { if ($_.Data['DriverName']) { $_.Data['DriverName'] } elseif ($_.Props.Count) { $_.Props[0] } } | Where-Object { $_ } | Select-Object -Unique); if ($d) { "Devices ($($d.Count)): " + ((@($d | Select-Object -First 6)) -join '; ') } } }

    # ---- services -----------------------------------------------------------------------------
    @{ Code = 'services.failed'; Log = 'System'; Provider = 'Service Control Manager'; Id = 7000, 7001, 7009, 7011, 7023, 7024, 7031, 7034; Severity = 'Warning'; MinCount = 2
        Title = 'Service keeps failing'; Detail = 'A service failed to start or stopped unexpectedly.'
        Group = { param($e) if ($e.Data['param1']) { $e.Data['param1'] } elseif ($e.Props.Count) { $e.Props[0] } else { 'service' } }
        Fix = { param($g) if ($g -match 'Windows Update|Background Intelligent|^wuauserv$|^BITS$') { 'reset-windows-update' } elseif ($g -match 'Print Spooler|^Spooler$') { 'clear-print-queue' } elseif ($g -match $script:AbCoreServiceNames) { 'start-core-services' } else { '' } } }

    # ---- applications ---------------------------------------------------------------------------
    @{ Code = 'apps.crash'; Log = 'Application'; Provider = '^Application Error$'; Id = 1000; Severity = 'Warning'; MinCount = 3
        Title = 'Application keeps crashing'; Detail = 'The app crashed repeatedly.'
        Group = { param($e) $a = if ($e.Props.Count) { $e.Props[0] } else { 'app' }; $m = if ($e.Props.Count -gt 3) { $e.Props[3] } else { '' }; "$a|$m" }
        Fix = { param($g) if ($script:AbWindowsApps -contains ($g -split '\|')[0]) { 'repair-system-files' } else { 'update-apps' } }
        Explain = { param($g) $p = $g -split '\|'; if ($p[1]) { "$($p[0]) crashed in $($p[1])." } else { "$($p[0]) crashed." } } }
    @{ Code = 'apps.hang'; Log = 'Application'; Provider = '^Application Hang$'; Id = 1002; Severity = 'Warning'; MinCount = 3
        Title = 'Application keeps freezing'; Detail = 'The app stopped responding and was closed.'
        Group = { param($e) if ($e.Props.Count) { $e.Props[0] } else { 'app' } }
        Fix = { param($g) if ($script:AbWindowsApps -contains $g) { 'repair-system-files' } else { 'update-apps' } } }
    @{ Code = 'apps.dotnet-crash'; Log = 'Application'; Provider = '^\.NET Runtime$'; Id = 1026; Severity = 'Warning'; MinCount = 3
        Title = '.NET application crash'; Detail = 'A .NET application stopped with an unhandled exception.'; FixTask = 'update-apps' }
    @{ Code = 'profile.temporary'; Log = 'Application'; Provider = 'User Profile Service|User Profiles Service'; Id = 1511, 1515; Severity = 'Critical'
        Title = 'User signed in with a temporary profile'; Detail = 'Desktop files and settings appear missing; changes are lost at sign-out. Needs a technician (profile registry repair).' }

    # ---- updates --------------------------------------------------------------------------------
    @{ Code = 'updates.install-failed'; Log = 'System'; Provider = 'WindowsUpdateClient'; Id = 20; Severity = 'Warning'
        Title = 'Windows Update install failed'; Detail = 'An update failed to install.'
        Group = { param($e) $c = "$($e.Data['errorCode'])"; if (-not $c -and $e.Props.Count) { $c = "$($e.Props[0])" }; if ($c -match '^\d+$') { $c = '0x{0:X8}' -f [uint32][int64]$c }; if ($c) { $c.ToUpperInvariant() -replace '^0X', '0x' } else { 'unknown' } }
        Fix = { param($g) if ($script:AbUpdateErrors.ContainsKey($g)) { $script:AbUpdateErrors[$g][1] } else { 'reset-windows-update' } }
        Explain = { param($g) if ($script:AbUpdateErrors.ContainsKey($g)) { "Error ${g}: $($script:AbUpdateErrors[$g][0])" } else { "Error $g." } }
        Summarize = { param($events) $t = @($events | ForEach-Object { $_.Data['updateTitle'] } | Where-Object { $_ } | Select-Object -Unique); if ($t) { 'Updates: ' + ((@($t | Select-Object -First 3)) -join '; ') } } }
    @{ Code = 'updates.servicing-error'; Log = 'Setup'; Provider = 'Servicing|CBS'; Id = 3, 4; Level = 1, 2; Severity = 'Warning'
        Title = 'Windows servicing error'; Detail = 'Installing a Windows package failed (component store).'; FixTask = 'repair-system-files' }

    # ---- network, time, domain ---------------------------------------------------------------
    @{ Code = 'network.dns-timeouts'; Log = 'System'; Provider = 'DNS Client Events|DNS-Client'; Id = 1014; Severity = 'Warning'; MinCount = 10
        Title = 'DNS lookups timing out'; Detail = 'Name resolution is failing intermittently (slow web pages, Outlook disconnects).'; FixTask = 'reset-network-cache' }
    @{ Code = 'time.sync-failed'; Log = 'System'; Provider = 'Time-Service'; Id = 129, 134, 36; Severity = 'Warning'; MinCount = 2
        Title = 'Clock could not sync'; Detail = 'The time service could not reach a time source.'; FixTask = 'resync-time' }
    @{ Code = 'domain.no-dc'; Log = 'System'; Provider = 'NETLOGON'; Id = 5719, 3210; Severity = 'Warning'
        Title = 'Cannot reach a domain controller'; Detail = 'Sign-in, Group Policy and drive mappings may fail. Check VPN/network and DNS.' }
    @{ Code = 'domain.gpo-failed'; Log = 'System'; Provider = 'GroupPolicy'; Id = 1129, 1055, 1053, 1058; Severity = 'Warning'; MinCount = 2
        Title = 'Group Policy failed'; Detail = 'Policies are not applying (network, DNS or domain trust).' }

    # ---- security -------------------------------------------------------------------------------
    @{ Code = 'security.malware-detected'; Log = 'Microsoft-Windows-Windows Defender/Operational'; Provider = 'Windows Defender'; Id = 1116, 1118, 1119; Severity = 'Critical'
        Title = 'Malware detected'; Detail = 'Defender detected a threat. 1118/1119 mean the clean-up failed. Investigate now.'
        Group = { param($e) if ($e.Data['Threat Name']) { $e.Data['Threat Name'] } else { 'threat' } } }

    # ---- slow start-up culprits (Diagnostics-Performance) --------------------------------------
    @{ Code = 'performance.startup-culprit'; Log = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Provider = 'Diagnostics-Performance'; Id = 101, 102, 103, 109; Severity = 'Info'; MinCount = 2
        Title = 'Slows down start-up'; Detail = 'Windows measured this app, driver or service delaying boot.'
        Group = { param($e) foreach ($k in 'FriendlyName', 'Name', 'FileName') { if ($e.Data[$k]) { return $e.Data[$k] } }; 'item' }
        Fix = { param($g) if ($g -match '\.sys$|driver') { 'install-driver-updates' } else { '' } } }
)

# Event sources that are noisy and harmless; excluded from the "many errors" count.
$script:AbEventNoise = @(
    @{ Provider = 'DistributedCOM'; Id = 10016, 10010 }
    @{ Provider = 'Schannel'; Id = 36887, 36874 }
    @{ Provider = 'Microsoft-Windows-Security-SPP'; Id = 8198 }
)

function Get-AbEventRule { $script:AbEventRules }

function Test-AbEventNoise {
    param($Record)
    foreach ($n in $script:AbEventNoise) { if ($Record.Provider -match $n.Provider -and $n.Id -contains $Record.Id) { return $true } }
    $false
}

function Invoke-AbEventAnalysis {
    <#
      Runs every rule over events from the last $Days days and returns findings. Events are read with
      one query per log (ids batched; Get-WinEvent rejects more than ~22 ids per filter).
    #>
    param([int]$Days = 7, [int]$MaxGroupsPerRule = 5, [object[]]$Rules = $script:AbEventRules)
    $since = (Get-AbDate).AddDays(-$Days)
    $byLog = @{}
    foreach ($r in $Rules) { if (-not $byLog.ContainsKey($r.Log)) { $byLog[$r.Log] = @() }; $byLog[$r.Log] += @($r.Id) }
    $events = @{}
    foreach ($log in $byLog.Keys) {
        $ids = @($byLog[$log] | Sort-Object -Unique)
        $list = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $ids.Count; $i += 20) {
            foreach ($e in @(Get-AbEventRecord -LogName $log -Id $ids[$i..([Math]::Min($i + 19, $ids.Count - 1))] -StartTime $since)) { if ($e) { $list.Add($e) } }
        }
        $events[$log] = $list
    }
    foreach ($r in $Rules) {
        $levels = if ($r.ContainsKey('Level')) { @($r.Level) } else { @(1, 2, 3) }
        $hits = @($events[$r.Log] | Where-Object {
                $_.Provider -match $r.Provider -and $r.Id -contains $_.Id -and $levels -contains $_.Level
            })
        if (-not $hits) { continue }
        $min = if ($r.ContainsKey('MinCount')) { [int]$r.MinCount } else { 1 }
        $groups = if ($r.ContainsKey('Group')) {
            @($hits | Group-Object -Property @{ Expression = { [string](& $r.Group $_) } } | Sort-Object Count -Descending)
        } else { @([pscustomobject]@{ Name = ''; Count = $hits.Count; Group = $hits }) }
        $shown = 0
        foreach ($g in $groups) {
            if ($g.Count -lt $min -or $shown -ge $MaxGroupsPerRule) { continue }
            $shown++
            $last = ($g.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
            $fix = if ($r.ContainsKey('Fix')) { [string](& $r.Fix $g.Name) } elseif ($r.ContainsKey('FixTask')) { $r.FixTask } else { '' }
            $explain = if ($r.ContainsKey('Explain')) { [string](& $r.Explain $g.Name) } else { '' }
            $subject = if ($g.Name -and -not $r.ContainsKey('Explain')) { ": $($g.Name)" } else { '' }
            $when = if ($last) { ' Last seen ' + ([datetime]$last).ToString('yyyy-MM-dd HH:mm') + '.' } else { '' }
            $summary = if ($r.ContainsKey('Summarize')) { [string](& $r.Summarize $g.Group) } else { '' }
            $detail = (@($explain, $r.Detail, $summary) | Where-Object { $_ }) -join ' '
            New-AbFinding -Code $r.Code -Severity $r.Severity -Title "$($r.Title)$subject ($($g.Count)x in $Days days)" -Detail "$detail$when" -FixTask $fix `
                -Data @{ Count = $g.Count; Group = $g.Name; Log = $r.Log; EventIds = (@($g.Group | ForEach-Object Id | Sort-Object -Unique) -join ',') }
        }
    }
}
