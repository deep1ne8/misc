#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }
# Behaviour of every check and fix, with all system access mocked (runs on any OS, changes nothing).

BeforeAll {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
    $env:AUTOBYTE_LOG_DIR = Join-Path ([IO.Path]::GetTempPath()) "ab-test-logs-$PID"
    Import-Module (Join-Path $root 'src/AutoByte/AutoByte.psd1') -Force
    $script:Now = [datetime]'2026-11-16T12:00:00'
    function script:Svc($Name, $Status = 'Running', $StartType = 'Automatic') {
        [pscustomobject]@{ Name = $Name; DisplayName = "$Name service"; Status = $Status; StartType = $StartType }
    }
    function script:File($Name, $AgeDays, $Length = 1MB) {
        [pscustomobject]@{ FullName = "/tmp/x/$Name"; Name = $Name; Length = $Length; LastWriteTime = $script:Now.AddDays(-$AgeDays) }
    }
}
AfterAll { Remove-Module AutoByte -ErrorAction SilentlyContinue }

Describe 'AutoByte tasks' {
BeforeEach {
    Mock -ModuleName AutoByte Get-AbDate { $script:Now }
    Mock -ModuleName AutoByte Test-AbIsAdmin { $true }
    Mock -ModuleName AutoByte Set-AbServiceState { }
    Mock -ModuleName AutoByte Set-AbServiceStartType { }
    Mock -ModuleName AutoByte Rename-AbItem { }
    Mock -ModuleName AutoByte Remove-AbFile { $true }
    Mock -ModuleName AutoByte Invoke-AbNative { [pscustomobject]@{ ExitCode = 0; Output = '' } }
}

Describe 'disk-space' {
    It 'is <expect> at <pct>% free on the system drive' -ForEach @(
        @{ pct = 3; expect = 'disk.critical' }, @{ pct = 10; expect = 'disk.low' }, @{ pct = 50; expect = '' }
    ) {
        $env:SystemDrive = 'C:'
        Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ DeviceID = 'C:'; Size = 100GB; FreeSpace = $pct * 1GB } } -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' }
        $f = (Invoke-AbTask -Id disk-space).Findings
        if ($expect) { $f.Code | Should -Be $expect; $f.FixTask | Should -Be 'clean-temp-files' } else { $f | Should -BeNullOrEmpty }
    }
    It 'does not suggest the temp clean-up for a full data drive' {
        $env:SystemDrive = 'C:'
        Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ DeviceID = 'D:'; Size = 100GB; FreeSpace = 1GB } } -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' }
        (Invoke-AbTask -Id disk-space).Findings.FixTask | Should -Be ''
    }
}

Describe 'disk-health / memory / uptime' {
    It 'flags an unhealthy physical disk as critical' {
        Mock -ModuleName AutoByte Get-AbPhysicalDisk { [pscustomobject]@{ FriendlyName = 'SSD'; HealthStatus = 'Warning'; OperationalStatus = 'Degraded' } }
        (Invoke-AbTask -Id disk-health).Findings.Severity | Should -Be 'Critical'
    }
    It 'flags low free memory' {
        Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ TotalVisibleMemorySize = 1000; FreePhysicalMemory = 50 } } -ParameterFilter { $ClassName -eq 'Win32_OperatingSystem' }
        (Invoke-AbTask -Id memory).Findings.Code | Should -Be 'system.memory-low'
    }
    It 'flags a PC not restarted for 20 days' {
        Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ LastBootUpTime = $script:Now.AddDays(-20) } } -ParameterFilter { $ClassName -eq 'Win32_OperatingSystem' }
        (Invoke-AbTask -Id uptime).Findings.Code | Should -Be 'system.uptime-long'
    }
}

Describe 'pending-reboot' {
    It 'reports which component wants a restart' {
        Mock -ModuleName AutoByte Test-AbPath { $Path -like '*WindowsUpdate*RebootRequired' }
        Mock -ModuleName AutoByte Get-AbRegistryValue { $null }
        $f = (Invoke-AbTask -Id pending-reboot).Findings
        $f.Code | Should -Be 'system.reboot-pending'
        $f.Detail | Should -BeLike '*Windows Update*'
    }
    It 'is quiet when nothing is pending' {
        Mock -ModuleName AutoByte Test-AbPath { $false }
        Mock -ModuleName AutoByte Get-AbRegistryValue { $null }
        (Invoke-AbTask -Id pending-reboot).Status | Should -Be 'Healthy'
    }
}

Describe 'core-services and start-core-services' {
    It 'finds stopped automatic services and the fix starts only those' {
        Mock -ModuleName AutoByte Get-AbService { @((Svc 'Dnscache' 'Stopped'), (Svc 'Spooler' 'Running'), (Svc 'AudioSrv' 'Stopped' 'Manual')) }
        (Invoke-AbTask -Id core-services).Findings.Data.Services | Should -Be @('Dnscache')
        $r = Invoke-AbTask -Id start-core-services -Confirm:$false
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -Times 1 -Exactly -ParameterFilter { $Name -eq 'Dnscache' -and $State -eq 'Running' }
        Should -Invoke -ModuleName AutoByte Set-AbServiceStartType -Times 0
        $r.Status | Should -Be 'Failed' -Because 'the mocked service never changes state, so verification correctly fails'
    }
}

Describe 'defender and update-defender-signatures' {
    It 'is critical when real-time protection is off' {
        Mock -ModuleName AutoByte Get-AbDefenderStatus { [pscustomobject]@{ AntivirusEnabled = $true; RealTimeProtectionEnabled = $false; AntivirusSignatureAge = 0 } }
        (Invoke-AbTask -Id defender).Findings.Code | Should -Be 'defender.realtime-off'
    }
    It 'suggests a definitions update when they are old' {
        Mock -ModuleName AutoByte Get-AbDefenderStatus { [pscustomobject]@{ AntivirusEnabled = $true; RealTimeProtectionEnabled = $true; AntivirusSignatureAge = 9 } }
        (Invoke-AbTask -Id defender).Findings.FixTask | Should -Be 'update-defender-signatures'
    }
    It 'is informational when another antivirus is in charge' {
        Mock -ModuleName AutoByte Get-AbDefenderStatus { throw 'not available' }
        (Invoke-AbTask -Id defender).Findings.Severity | Should -Be 'Info'
    }
    It 'fails the fix when MpCmdRun fails' {
        Mock -ModuleName AutoByte Get-AbDefenderStatus { [pscustomobject]@{ AntivirusSignatureAge = 9 } }
        Mock -ModuleName AutoByte Invoke-AbNative { [pscustomobject]@{ ExitCode = 2; Output = 'x' } }
        (Invoke-AbTask -Id update-defender-signatures -Confirm:$false).Status | Should -Be 'Failed'
    }
}

Describe 'connectivity and reset-network-cache' {
    It 'is critical when no name resolves' {
        Mock -ModuleName AutoByte Resolve-AbDns { $false }
        $f = (Invoke-AbTask -Id connectivity).Findings
        $f.Code | Should -Be 'network.dns-down'
        $f.FixTask | Should -Be 'reset-network-cache'
    }
    It 'is critical (offline) when names resolve but HTTPS is blocked' {
        Mock -ModuleName AutoByte Resolve-AbDns { $true }
        Mock -ModuleName AutoByte Test-AbTcp { $false }
        (Invoke-AbTask -Id connectivity).Findings.Code | Should -Be 'network.offline'
    }
    It 'flushes DNS when lookups fail and does not report failure if the network stays down' {
        Mock -ModuleName AutoByte Resolve-AbDns { $false }
        $r = Invoke-AbTask -Id reset-network-cache -Confirm:$false
        $r.Status | Should -Be 'Fixed'
        Should -Invoke -ModuleName AutoByte Invoke-AbNative -Times 1 -ParameterFilter { $ArgumentList -contains '/flushdns' }
        ($r.Messages -join ' ') | Should -BeLike '*Still failing*'
    }
    It 'does nothing when DNS works (unless -Force)' {
        Mock -ModuleName AutoByte Resolve-AbDns { $true }
        (Invoke-AbTask -Id reset-network-cache -Confirm:$false).Status | Should -Be 'NotNeeded'
        (Invoke-AbTask -Id reset-network-cache -Parameters @{ Force = $true } -Confirm:$false).Status | Should -Be 'Fixed'
    }
}

Describe 'time-sync and resync-time' {
    It 'is critical when the clock is 400 s off and suggests a resync' {
        Mock -ModuleName AutoByte Get-AbTimeOffset { 400 }
        $f = (Invoke-AbTask -Id time-sync).Findings
        $f.Severity | Should -Be 'Critical'
        $f.FixTask | Should -Be 'resync-time'
    }
    It 'never changes the clock when the offset cannot be measured' {
        Mock -ModuleName AutoByte Get-AbTimeOffset { $null }
        (Invoke-AbTask -Id resync-time -Confirm:$false).Status | Should -Be 'NotNeeded'
        Should -Invoke -ModuleName AutoByte Invoke-AbNative -Times 0
    }
    It 'starts W32Time and resyncs' {
        $script:offsets = [System.Collections.Generic.Queue[object]]::new(); $script:offsets.Enqueue(120); $script:offsets.Enqueue(0.2)
        Mock -ModuleName AutoByte Get-AbTimeOffset { $script:offsets.Dequeue() }
        Mock -ModuleName AutoByte Get-AbService { Svc 'W32Time' 'Stopped' 'Manual' }
        (Invoke-AbTask -Id resync-time -Confirm:$false).Status | Should -Be 'Fixed'
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -ParameterFilter { $Name -eq 'W32Time' -and $State -eq 'Running' }
        Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains '/resync' }
    }
}

Describe 'windows-update and reset-windows-update' {
    It 'is critical when wuauserv is disabled' {
        Mock -ModuleName AutoByte Get-AbService { Svc 'wuauserv' 'Stopped' 'Disabled' }
        Mock -ModuleName AutoByte Get-AbCim { } -ParameterFilter { $ClassName -eq 'Win32_QuickFixEngineering' }
        (Invoke-AbTask -Id windows-update).Findings.Code | Should -Be 'updates.service-disabled'
    }
    It 'warns when no update installed for 60 days' {
        Mock -ModuleName AutoByte Get-AbService { Svc 'wuauserv' 'Running' 'Manual' }
        Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ InstalledOn = $script:Now.AddDays(-60) } } -ParameterFilter { $ClassName -eq 'Win32_QuickFixEngineering' }
        (Invoke-AbTask -Id windows-update).Findings.Code | Should -Be 'updates.stale'
    }
    It 'reset re-enables, renames caches, restarts services and asks for a restart' {
        $script:disabled = $true
        Mock -ModuleName AutoByte Get-AbService { foreach ($n in $Name) { if ($n -eq 'wuauserv' -and $script:disabled) { Svc $n 'Stopped' 'Disabled' } else { Svc $n 'Running' 'Manual' } } }
        Mock -ModuleName AutoByte Set-AbServiceStartType { $script:disabled = $false }
        Mock -ModuleName AutoByte Test-AbPath { $true }
        $r = Invoke-AbTask -Id reset-windows-update -Confirm:$false
        $r.Status | Should -Be 'Fixed'
        $r.RebootRequired | Should -BeTrue
        Should -Invoke -ModuleName AutoByte Set-AbServiceStartType -ParameterFilter { $Name -eq 'wuauserv' -and $StartupType -eq 'Manual' }
        Should -Invoke -ModuleName AutoByte Rename-AbItem -Times 2 -Exactly
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -ParameterFilter { $Name -eq 'wuauserv' -and $State -eq 'Running' }
    }
    It 'reset is not needed when services are enabled, unless forced' {
        Mock -ModuleName AutoByte Get-AbService { foreach ($n in $Name) { Svc $n 'Running' 'Manual' } }
        Mock -ModuleName AutoByte Test-AbPath { $true }
        (Invoke-AbTask -Id reset-windows-update -Confirm:$false).Status | Should -Be 'NotNeeded'
        (Invoke-AbTask -Id reset-windows-update -Parameters @{ Force = $true } -Confirm:$false).Status | Should -Be 'Fixed'
    }
}

Describe 'repair-system-files' {
    It 'runs DISM then SFC when the store is repairable' {
        $script:health = [System.Collections.Generic.Queue[object]]::new(); $script:health.Enqueue('Repairable'); $script:health.Enqueue('Healthy')
        Mock -ModuleName AutoByte Get-AbImageHealth { $script:health.Dequeue() }
        Mock -ModuleName AutoByte Repair-AbImage { [pscustomobject]@{ State = 'Healthy'; RestartNeeded = $false } }
        (Invoke-AbTask -Id repair-system-files -Confirm:$false).Status | Should -Be 'Fixed'
        Should -Invoke -ModuleName AutoByte Repair-AbImage -Times 1
        Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains '/scannow' }
    }
    It 'fails clearly when the image is not repairable' {
        Mock -ModuleName AutoByte Get-AbImageHealth { 'NonRepairable' }
        $r = Invoke-AbTask -Id repair-system-files -Confirm:$false
        $r.Status | Should -Be 'Failed'
        $r.Error | Should -BeLike '*in-place upgrade*'
    }
}

Describe 'temp-files and clean-temp-files' {
    BeforeEach {
        Mock -ModuleName AutoByte Get-AbSystemPath { "/tmp/$Name" }
        Mock -ModuleName AutoByte Get-AbChildItem { if ($Path -eq '/tmp/UserTemp') { 1..300 | ForEach-Object { File "old$_" 30 10MB } ; File 'new' 1 } }
    }
    It 'reports the reclaimable space' {
        (Invoke-AbTask -Id temp-files).Findings.FixTask | Should -Be 'clean-temp-files'
    }
    It 'deletes only old files and verifies even if some are in use' {
        Mock -ModuleName AutoByte Remove-AbFile { $Path -notlike '*old1' }
        $r = Invoke-AbTask -Id clean-temp-files -Confirm:$false
        $r.Status | Should -Be 'Fixed'
        Should -Invoke -ModuleName AutoByte Remove-AbFile -Times 300 -Exactly
        Should -Invoke -ModuleName AutoByte Remove-AbFile -Times 0 -ParameterFilter { $Path -like '*new' }
    }
    It 'deletes nothing under -WhatIf' {
        (Invoke-AbTask -Id clean-temp-files -WhatIf).Status | Should -Be 'WhatIf'
        Should -Invoke -ModuleName AutoByte Remove-AbFile -Times 0
    }
    It 'is not needed below the threshold' {
        Mock -ModuleName AutoByte Get-AbChildItem { File 'old' 30 }
        (Invoke-AbTask -Id clean-temp-files -Confirm:$false).Status | Should -Be 'NotNeeded'
    }
}

Describe 'print-queue and clear-print-queue' {
    BeforeEach { Mock -ModuleName AutoByte Get-AbSystemPath { "/tmp/$Name" } }
    It 'detects stuck jobs' {
        Mock -ModuleName AutoByte Get-AbService { Svc 'Spooler' }
        Mock -ModuleName AutoByte Get-AbChildItem { (File 'a.SHD' 0.1), (File 'a.SPL' 0.1) }
        (Invoke-AbTask -Id print-queue).Findings.Code | Should -Be 'print.stuck-jobs'
    }
    It 'stops the spooler, removes old jobs, and always starts it again' {
        $script:files = @((File 'a.SHD' 0.1), (File 'a.SPL' 0.1))
        Mock -ModuleName AutoByte Get-AbService { Svc 'Spooler' }
        Mock -ModuleName AutoByte Get-AbChildItem { $script:files }
        Mock -ModuleName AutoByte Remove-AbFile { $script:files = @(); $true }
        (Invoke-AbTask -Id clear-print-queue -Confirm:$false).Status | Should -Be 'Fixed'
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -ParameterFilter { $State -eq 'Stopped' } -Times 1 -Exactly
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -ParameterFilter { $State -eq 'Running' } -Times 1 -Exactly
    }
    It 'restarts the spooler even if deleting fails' {
        Mock -ModuleName AutoByte Get-AbService { Svc 'Spooler' }
        Mock -ModuleName AutoByte Get-AbChildItem { File 'a.SPL' 0.1 }
        Mock -ModuleName AutoByte Remove-AbFile { throw 'disk error' }
        (Invoke-AbTask -Id clear-print-queue -Confirm:$false).Status | Should -Be 'Failed'
        Should -Invoke -ModuleName AutoByte Set-AbServiceState -ParameterFilter { $State -eq 'Running' } -Times 1 -Exactly
    }
}

Describe 'office' {
    It 'reports version and channel, and warns when updates are off' {
        Mock -ModuleName AutoByte Get-AbRegistryValue {
            switch ($Name) { 'VersionToReport' { '16.0.18000.20000' } 'Platform' { 'x64' } 'UpdateChannel' { 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60' } 'UpdatesEnabled' { 'False' } default { $null } }
        }
        $f = (Invoke-AbTask -Id office).Findings
        ($f | Where-Object Code -EQ 'office.info').Data.Channel | Should -Be 'Current'
        $f.Code | Should -Contain 'office.updates-off'
    }
    It 'is silent when Click-to-Run Office is not installed' {
        Mock -ModuleName AutoByte Get-AbRegistryValue { $null }
        (Invoke-AbTask -Id office).Findings | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-AbScan -> Invoke-AbFix' {
    It 'runs only the fixes suggested by non-info findings' {
        $scan = [pscustomobject]@{ SuggestedFixes = @('reset-network-cache') }
        Mock -ModuleName AutoByte Resolve-AbDns { $false }
        $r = @($scan | Invoke-AbFix -Confirm:$false)
        $r.TaskId | Should -Be 'reset-network-cache'
    }
}
}
