#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }
# Event-log analysis, performance, drivers, updates, CVE exposure and app updates. All system access mocked.

BeforeAll {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
    $env:AUTOBYTE_LOG_DIR = Join-Path ([IO.Path]::GetTempPath()) "ab-test-logs-$PID"
    Import-Module (Join-Path $root 'src/AutoByte/AutoByte.psd1') -Force
    $script:Now = [datetime]'2026-11-16T12:00:00'
    function script:Ev($Provider, $Id, [string[]]$Props = @(), [hashtable]$Data = @{}, $Level = 2, $Log = 'System') {
        [pscustomobject]@{ Log = $Log; Provider = $Provider; Id = $Id; Level = $Level; Time = $script:Now.AddHours(-1); Data = $Data; Props = $Props }
    }
    function script:Upd($Title, $Severity = '', [string[]]$Cat = @('0fa1201d-4330-4fa8-8ae9-b877473b6441')) {
        [pscustomobject]@{ UpdateId = [guid]::NewGuid().ToString(); Title = $Title; KB = 'KB5000001'; Severity = $Severity; CategoryIds = $Cat; Type = 1; SizeMB = 1; Reboot = $true }
    }
    # Minimal MSRC CVRF document (same shape as the real feed).
    $script:CvrfJson = @'
{ "ProductTree": { "Branch": [ { "Items": [ { "Items": [
      { "ProductID": "100", "Value": "Windows 11 Version 24H2 for x64-based Systems" },
      { "ProductID": "200", "Value": "Windows 10 Version 22H2 for x64-based Systems" } ] } ] } ],
    "FullProductName": [ { "ProductID": "100", "Value": "Windows 11 Version 24H2 for x64-based Systems" },
                         { "ProductID": "200", "Value": "Windows 10 Version 22H2 for x64-based Systems" } ] },
  "Vulnerability": [
    { "CVE": "CVE-2026-1001", "Title": { "Value": "Windows Kernel Elevation of Privilege" },
      "Threats": [ { "Type": 3, "Description": { "Value": "Important" }, "ProductID": [ "100" ] },
                   { "Type": 1, "Description": { "Value": "Publicly Disclosed:No;Exploited:Yes;Latest Software Release:Exploitation Detected" } } ],
      "CVSSScoreSets": [ { "BaseScore": 7.8, "ProductID": [ "100" ] } ],
      "Remediations": [ { "Type": 2, "Description": { "Value": "5050001" }, "ProductID": [ "100" ], "FixedBuild": "10.0.26100.3000" },
                        { "Type": 2, "Description": { "Value": "5050002" }, "ProductID": [ "200" ], "FixedBuild": "10.0.19045.5500" } ] },
    { "CVE": "CVE-2026-1002", "Title": { "Value": "Windows TCP/IP Remote Code Execution" },
      "Threats": [ { "Type": 3, "Description": { "Value": "Critical" }, "ProductID": [ "100" ] } ],
      "Remediations": [ { "Type": 2, "Description": { "Value": "5050001" }, "ProductID": [ "100" ], "FixedBuild": "10.0.26100.2800" } ] },
    { "CVE": "CVE-2026-1003", "Title": { "Value": "Old bug" },
      "Remediations": [ { "Type": 2, "Description": { "Value": "5040000" }, "ProductID": [ "100" ], "FixedBuild": "10.0.26100.2000" } ] },
    { "CVE": "CVE-2026-1004", "Title": { "Value": "Windows 10 only" },
      "Remediations": [ { "Type": 2, "Description": { "Value": "5050002" }, "ProductID": [ "200" ], "FixedBuild": "10.0.19045.5500" } ] }
  ] }
'@
    $script:Os = [pscustomobject]@{ Build = '26100'; Ubr = 2500; DisplayVersion = '24H2'; ProductName = 'Windows 11 Pro'; InstallationType = 'Client'; Architecture = 'AMD64' }
    function script:ConvertTo-Dict($o) {
        # Simulates Windows PowerShell 5.1's JavaScriptSerializer output (dictionaries + object[]).
        if ($o -is [System.Management.Automation.PSCustomObject]) { $d = New-Object 'System.Collections.Generic.Dictionary[string,object]'; foreach ($p in $o.PSObject.Properties) { $d[$p.Name] = ConvertTo-Dict $p.Value }; return , $d }
        if ($o -is [array]) { return , [object[]]@($o | ForEach-Object { ConvertTo-Dict $_ }) }
        $o
    }
}
AfterAll { Remove-Module AutoByte -ErrorAction SilentlyContinue }

Describe 'AutoByte diagnostics' {
    BeforeEach {
        Mock -ModuleName AutoByte Get-AbDate { $script:Now }
        Mock -ModuleName AutoByte Test-AbIsAdmin { $true }
        Mock -ModuleName AutoByte Wait-AbInterval { }
        Mock -ModuleName AutoByte Get-AbWinEvent { }
        Mock -ModuleName AutoByte Get-AbEventRecord { }
        Mock -ModuleName AutoByte Invoke-AbNative { [pscustomobject]@{ ExitCode = 0; Output = '' } }
    }

    Context 'event-logs' {
        It 'diagnoses a blue screen by stop code and maps it to the driver fix' {
            Mock -ModuleName AutoByte Get-AbEventRecord { Ev 'Microsoft-Windows-WER-SystemErrorReporting' 1001 @('0x0000009f (0x0000000000000003, 0xffff)', 'C:\Windows\MEMORY.DMP') } -ParameterFilter { $LogName -eq 'System' }
            $f = (Invoke-AbTask -Id event-logs).Findings | Where-Object Code -EQ 'events.bugcheck'
            $f.Severity | Should -Be 'Critical'
            $f.FixTask | Should -Be 'install-driver-updates'
            $f.Detail | Should -Match 'DRIVER_POWER_STATE_FAILURE'
        }
        It 'groups application crashes by app, honours MinCount, and picks the fix per app' {
            Mock -ModuleName AutoByte Get-AbEventRecord {
                1..3 | ForEach-Object { Ev 'Application Error' 1000 @('outlook.exe', '16.0', 'x', 'mso.dll') -Log Application }
                1..3 | ForEach-Object { Ev 'Application Error' 1000 @('explorer.exe', '10.0', 'x', 'shell32.dll') -Log Application }
                1..2 | ForEach-Object { Ev 'Application Error' 1000 @('notepad.exe', '1', 'x', 'y.dll') -Log Application }
            } -ParameterFilter { $LogName -eq 'Application' }
            $f = @((Invoke-AbTask -Id event-logs).Findings | Where-Object Code -EQ 'apps.crash')
            $f.Count | Should -Be 2
            ($f | Where-Object { $_.Data.Group -like 'outlook.exe*' }).FixTask | Should -Be 'update-apps'
            ($f | Where-Object { $_.Data.Group -like 'explorer.exe*' }).FixTask | Should -Be 'repair-system-files'
        }
        It 'maps Windows Update error codes (hex or decimal) to the matching fix' {
            Mock -ModuleName AutoByte Get-AbEventRecord {
                Ev 'Microsoft-Windows-WindowsUpdateClient' 20 -Data @{ errorCode = '0x800f081f'; updateTitle = 'CU' }
                Ev 'Microsoft-Windows-WindowsUpdateClient' 20 -Data @{ errorCode = '2147942512' }   # 0x80070070
            } -ParameterFilter { $LogName -eq 'System' }
            $f = @((Invoke-AbTask -Id event-logs).Findings | Where-Object Code -EQ 'updates.install-failed')
            ($f | Where-Object { $_.Data.Group -eq '0x800F081F' }).FixTask | Should -Be 'repair-system-files'
            ($f | Where-Object { $_.Data.Group -eq '0x80070070' }).FixTask | Should -Be 'clean-temp-files'
        }
        It 'routes a failing Windows Update service to the update reset' {
            Mock -ModuleName AutoByte Get-AbEventRecord { 1..2 | ForEach-Object { Ev 'Service Control Manager' 7031 -Data @{ param1 = 'Windows Update' } } } -ParameterFilter { $LogName -eq 'System' }
            ((Invoke-AbTask -Id event-logs).Findings | Where-Object Code -EQ 'services.failed').FixTask | Should -Be 'reset-windows-update'
        }
        It 'detects disk, driver, hardware and malware events' {
            Mock -ModuleName AutoByte Get-AbEventRecord { Ev 'disk' 7; Ev 'Microsoft-Windows-WHEA-Logger' 18; Ev 'Display' 4101 } -ParameterFilter { $LogName -eq 'System' }
            Mock -ModuleName AutoByte Get-AbEventRecord { Ev 'Microsoft-Windows-Windows Defender' 1116 -Data @{ 'Threat Name' = 'Trojan:Win32/Test' } } -ParameterFilter { $LogName -like '*Defender*' }
            $codes = (Invoke-AbTask -Id event-logs).Findings.Code
            foreach ($c in 'storage.bad-blocks', 'hardware.whea', 'drivers.display-reset', 'security.malware-detected') { $codes | Should -Contain $c }
        }
        It 'ignores informational events such as Ntfs 98 "volume is healthy"' {
            Mock -ModuleName AutoByte Get-AbEventRecord { 1..12 | ForEach-Object { Ev 'Ntfs' 98 -Level 4 } } -ParameterFilter { $LogName -eq 'System' }
            (Invoke-AbTask -Id event-logs).Findings | Should -BeNullOrEmpty
            Mock -ModuleName AutoByte Get-AbEventRecord { Ev 'Ntfs' 55 -Level 2 } -ParameterFilter { $LogName -eq 'System' }
            (Invoke-AbTask -Id event-logs).Findings.Code | Should -Be 'storage.filesystem-corrupt'
        }
        It 'summarises late-loading drivers as one note instead of a warning per device' {
            Mock -ModuleName AutoByte Get-AbEventRecord { foreach ($d in 'ACPI\INTC1073', 'USB\VID_04F3') { 1..3 | ForEach-Object { Ev 'Microsoft-Windows-Kernel-PnP' 219 -Level 3 -Data @{ DriverName = $d } } } } -ParameterFilter { $LogName -eq 'System' }
            $f = @((Invoke-AbTask -Id event-logs).Findings)
            $f.Count | Should -Be 1
            $f.Severity | Should -Be 'Info'
            $f.Detail | Should -Match 'Devices \(2\)'
        }
        It 'offers no fix for third-party services and Store app update errors' {
            Mock -ModuleName AutoByte Get-AbEventRecord {
                1..3 | ForEach-Object { Ev 'Service Control Manager' 7031 -Data @{ param1 = 'Cloudflared agent' } }
                1..2 | ForEach-Object { Ev 'Service Control Manager' 7000 -Data @{ param1 = 'DNS Client' } }
                Ev 'Microsoft-Windows-WindowsUpdateClient' 20 -Data @{ errorCode = '0x80073D02'; updateTitle = '9WZDNCRFJ3PT-Microsoft.ZuneMusic' }
            } -ParameterFilter { $LogName -eq 'System' }
            $f = (Invoke-AbTask -Id event-logs).Findings
            ($f | Where-Object { $_.Data.Group -eq 'Cloudflared agent' }).FixTask | Should -Be ''
            ($f | Where-Object { $_.Data.Group -eq 'DNS Client' }).FixTask | Should -Be 'start-core-services'
            $u = $f | Where-Object Code -EQ 'updates.install-failed'
            $u.FixTask | Should -Be ''
            $u.Detail | Should -Match 'Store app.*Updates: 9WZDNCRFJ3PT'
        }
        It 'queries each log with at most 20 event ids and the configured look-back' {
            Invoke-AbTask -Id event-logs -Parameters @{ Days = 3 } | Out-Null
            Should -Invoke -ModuleName AutoByte Get-AbEventRecord -ParameterFilter { $Id.Count -gt 20 } -Times 0 -Exactly
            Should -Invoke -ModuleName AutoByte Get-AbEventRecord -ParameterFilter { $StartTime -eq $script:Now.AddDays(-3) } -Times 1
        }
        It 'ignores known noise when counting errors' {
            Mock -ModuleName AutoByte Get-AbWinEvent { 1..40 | ForEach-Object { [pscustomobject]@{ ProviderName = 'Microsoft-Windows-DistributedCOM'; Id = 10016 } } }
            (Invoke-AbTask -Id event-logs).Findings | Should -BeNullOrEmpty
            Mock -ModuleName AutoByte Get-AbWinEvent { 1..40 | ForEach-Object { [pscustomobject]@{ ProviderName = 'Foo'; Id = 1 } } }
            (Invoke-AbTask -Id event-logs).Findings.Code | Should -Contain 'events.many-errors'
        }
    }

    Context 'performance and set-power-plan' {
        BeforeEach {
            Mock -ModuleName AutoByte Get-AbSystemDiskMediaType { 'SSD' }
            Mock -ModuleName AutoByte Get-AbCim { }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ NumberOfLogicalProcessors = 4 } } -ParameterFilter { $ClassName -eq 'Win32_ComputerSystem' }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentProcessorTime = 10 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfOS_Processor' }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentIdleTime = 90 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfDisk_PhysicalDisk' }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentCommittedBytesInUse = 40 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfOS_Memory' }
            Mock -ModuleName AutoByte Get-AbCim {
                [pscustomobject]@{ Name = 'MsMpEng'; PercentProcessorTime = 320; WorkingSetPrivate = 300MB; IODataBytesPersec = 5MB }
                [pscustomobject]@{ Name = 'chrome#3'; PercentProcessorTime = 40; WorkingSetPrivate = 900MB; IODataBytesPersec = 1MB }
                [pscustomobject]@{ Name = '_Total'; PercentProcessorTime = 400; WorkingSetPrivate = 1; IODataBytesPersec = 1 }
            } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfProc_Process' }
        }
        It 'is healthy on a quiet machine' {
            (Invoke-AbTask -Id performance).Status | Should -Be 'Healthy'
        }
        It 'names the process behind high CPU (per-core percentages normalised)' {
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentProcessorTime = 95 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfOS_Processor' }
            $f = (Invoke-AbTask -Id performance).Findings | Where-Object Code -EQ 'performance.cpu-high'
            $f.Data.Top | Should -Match '^MsMpEng \(80%\)'
            $f.Data.Top | Should -Not -Match '_Total'
        }
        It 'explains memory pressure, a busy disk, an HDD, power saver and a slow boot' {
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentCommittedBytesInUse = 95 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfOS_Memory' }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ PercentIdleTime = 2 } } -ParameterFilter { $ClassName -eq 'Win32_PerfFormattedData_PerfDisk_PhysicalDisk' }
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ IsActive = $true; InstanceID = 'Microsoft:PowerPlan\{a1841308-3541-4fab-bc81-f71556f20b4a}' } } -ParameterFilter { $ClassName -eq 'Win32_PowerPlan' }
            Mock -ModuleName AutoByte Get-AbSystemDiskMediaType { 'HDD' }
            Mock -ModuleName AutoByte Get-AbEventRecord { Ev 'Microsoft-Windows-Diagnostics-Performance' 100 -Data @{ BootTime = '185000' } } -ParameterFilter { $LogName -like '*Diagnostics-Performance*' }
            $f = (Invoke-AbTask -Id performance).Findings
            $f.Code | Should -Be @('performance.memory-pressure', 'performance.disk-busy', 'performance.hdd-system-disk', 'performance.power-saver', 'performance.slow-boot')
            ($f | Where-Object Code -EQ 'performance.memory-pressure').Data.Top | Should -Match '^chrome \(900 MB\)'
            ($f | Where-Object Code -EQ 'performance.power-saver').FixTask | Should -Be 'set-power-plan'
        }
        It 'switches to Balanced only when Power saver is active' {
            (Invoke-AbTask -Id set-power-plan -Confirm:$false).Status | Should -Be 'NotNeeded'
            Mock -ModuleName AutoByte Get-AbCim { [pscustomobject]@{ IsActive = $true; InstanceID = '{a1841308-3541-4fab-bc81-f71556f20b4a}' } } -ParameterFilter { $ClassName -eq 'Win32_PowerPlan' }
            Invoke-AbTask -Id set-power-plan -Confirm:$false | Out-Null
            Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains '381b4222-f694-41f0-9685-ff5bb260df2e' } -Times 1 -Exactly
        }
    }

    Context 'drivers' {
        BeforeEach {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { }
            Mock -ModuleName AutoByte Get-AbCim {
                [pscustomobject]@{ Name = 'Wi-Fi'; ConfigManagerErrorCode = 28; PNPClass = 'Net'; HardwareID = @('PCI\VEN_1'); PNPDeviceID = 'PCI\1' }
                [pscustomobject]@{ Name = 'Webcam'; ConfigManagerErrorCode = 22; PNPClass = 'Camera'; HardwareID = @('USB\1'); PNPDeviceID = 'USB\1' }
                [pscustomobject]@{ Name = 'Audio'; ConfigManagerErrorCode = 10; PNPClass = 'Media'; HardwareID = @('HDA\1'); PNPDeviceID = 'HDA\1' }
            } -ParameterFilter { $ClassName -eq 'Win32_PnPEntity' }
        }
        It 'reports problem devices with the Device Manager reason and skips disabled ones' {
            $f = @((Invoke-AbTask -Id drivers).Findings | Where-Object Code -EQ 'drivers.device-problem')
            $f.Count | Should -Be 2
            ($f | Where-Object { $_.Data.Device -eq 'Wi-Fi' }).FixTask | Should -Be 'install-driver-updates'
            ($f | Where-Object { $_.Data.Device -eq 'Audio' }).FixTask | Should -Be 'rescan-devices'
        }
        It 'offers driver updates and fails soft when Windows Update cannot be searched' {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { Upd 'Intel - Net - 23.1' }
            (Invoke-AbTask -Id drivers).Findings.Code | Should -Contain 'drivers.updates-available'
            Mock -ModuleName AutoByte Get-AbPendingUpdate { throw 'COM error' }
            $r = Invoke-AbTask -Id drivers
            $r.Status | Should -Be 'Issues'
            $r.Findings.Code | Should -Contain 'drivers.update-search-failed'
        }
        It 'rescans hardware with pnputil' {
            (Invoke-AbTask -Id rescan-devices -Confirm:$false).Status | Should -Be 'Failed' -Because 'mocked devices stay broken, so verification correctly fails'
            Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains '/scan-devices' } -Times 1 -Exactly
        }
        It 'installs driver updates and reports the restart' {
            $script:drvInstalled = $false
            Mock -ModuleName AutoByte Get-AbPendingUpdate { if (-not $script:drvInstalled) { Upd 'Realtek - Audio' } }
            Mock -ModuleName AutoByte Install-AbUpdate { $script:drvInstalled = $true; [pscustomobject]@{ Installed = 1; Failed = 0; RebootRequired = $true; Results = @() } }
            $r = Invoke-AbTask -Id install-driver-updates -Confirm:$false
            $r.Status | Should -Be 'Fixed'
            $r.RebootRequired | Should -BeTrue
            Should -Invoke -ModuleName AutoByte Install-AbUpdate -ParameterFilter { $Criteria -match "Type='Driver'" } -Times 1 -Exactly
        }
    }

    Context 'security updates' {
        It 'flags missing critical security updates' {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { Upd '2026-11 Cumulative Update' 'Critical'; Upd 'Feature' '' @('other') }
            $f = (Invoke-AbTask -Id security-updates).Findings
            ($f | Where-Object Code -EQ 'updates.security-missing').Severity | Should -Be 'Critical'
            $f.Code | Should -Contain 'updates.other-available'
        }
        It 'turns a search failure into a finding with the reset fix' {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { throw '0x8024402C' }
            (Invoke-AbTask -Id security-updates).Findings.FixTask | Should -Be 'reset-windows-update'
        }
        It 'installs only security updates unless -All, and verifies' {
            $script:secInstalled = $false
            Mock -ModuleName AutoByte Get-AbPendingUpdate { if (-not $script:secInstalled) { [pscustomobject]@{ UpdateId = 'sec'; Title = 'CU'; Severity = 'Important'; CategoryIds = @(); KB = '' } }; [pscustomobject]@{ UpdateId = 'opt'; Title = 'Optional'; Severity = ''; CategoryIds = @('x'); KB = '' } }
            Mock -ModuleName AutoByte Install-AbUpdate { $script:secInstalled = $true; [pscustomobject]@{ Installed = 1; Failed = 0; RebootRequired = $true; Results = @() } }
            (Invoke-AbTask -Id install-security-updates -Confirm:$false).Status | Should -Be 'Fixed'
            Should -Invoke -ModuleName AutoByte Install-AbUpdate -ParameterFilter { $UpdateId -contains 'sec' -and $UpdateId -notcontains 'opt' } -Times 1 -Exactly
        }
        It 'fails when every update fails' {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { Upd 'CU' 'Critical' }
            Mock -ModuleName AutoByte Install-AbUpdate { [pscustomobject]@{ Installed = 0; Failed = 1; RebootRequired = $false; Results = @() } }
            (Invoke-AbTask -Id install-security-updates -Confirm:$false).Status | Should -Be 'Failed'
        }
        It 'supports -WhatIf without installing' {
            Mock -ModuleName AutoByte Get-AbPendingUpdate { Upd 'CU' 'Critical' }
            Mock -ModuleName AutoByte Install-AbUpdate { }
            (Invoke-AbTask -Id install-security-updates -WhatIf).Status | Should -Be 'WhatIf'
            Should -Invoke -ModuleName AutoByte Install-AbUpdate -Times 0 -Exactly
        }
    }

    Context 'cve-exposure' {
        BeforeEach {
            Mock -ModuleName AutoByte Get-AbOsBuild { $script:Os }
            Mock -ModuleName AutoByte Get-AbCvrfDocument { , ($script:CvrfJson | ConvertFrom-Json) }
            Mock -ModuleName AutoByte Get-AbKevCve { 'CVE-2026-1001' }
        }
        It 'maps the OS to MSRC product names' {
            InModuleScope AutoByte {
                Get-AbMsrcProductName -Os ([pscustomobject]@{ Build = '26100'; DisplayVersion = '24H2'; InstallationType = 'Client'; Architecture = 'AMD64' }) | Should -Be 'Windows 11 Version 24H2 for x64-based Systems'
                Get-AbMsrcProductName -Os ([pscustomobject]@{ Build = '19045'; DisplayVersion = '22H2'; InstallationType = 'Client'; Architecture = 'ARM64' }) | Should -Be 'Windows 10 Version 22H2 for ARM64-based Systems'
                Get-AbMsrcProductName -Os ([pscustomobject]@{ Build = '20348'; DisplayVersion = ''; InstallationType = 'Server Core'; Architecture = 'AMD64' }) | Should -Be 'Windows Server 2022 (Server Core installation)'
            }
        }
        It 'lists CVEs fixed in newer builds, flags KEV ones critical, and ignores other OS branches' {
            $f = (Invoke-AbTask -Id cve-exposure).Findings
            $f.Code | Should -Be 'cve.missing-fixes'
            $f.Severity | Should -Be 'Critical'
            $f.FixTask | Should -Be 'install-security-updates'
            $f.Data.Cves -split ',' | Sort-Object | Should -Be @('CVE-2026-1001', 'CVE-2026-1002')
            $f.Data.RequiredBuild | Should -Be '10.0.26100.3000'
            $f.Data.Exploited | Should -Be 1
            $f.Detail | Should -Match 'CISA KEV'
        }
        It 'reports up to date when the build is current' {
            Mock -ModuleName AutoByte Get-AbOsBuild { $o = $script:Os.PSObject.Copy(); $o.Ubr = 3000; $o }
            (Invoke-AbTask -Id cve-exposure).Findings.Code | Should -Be 'cve.up-to-date'
        }
        It 'does not claim "patched" for a build the feed does not cover (Insider)' {
            Mock -ModuleName AutoByte Get-AbOsBuild { $o = $script:Os.PSObject.Copy(); $o.Build = '26300'; $o.Ubr = 9974; $o }
            $f = (Invoke-AbTask -Id cve-exposure).Findings
            $f.Code | Should -Be 'cve.product-not-in-feed'
            $f.Detail | Should -Match 'Windows 11 Version 24H2 for x64-based Systems'
        }
        It 'names the full product in the up-to-date note' {
            Mock -ModuleName AutoByte Get-AbOsBuild { $o = $script:Os.PSObject.Copy(); $o.Ubr = 3000; $o }
            (Invoke-AbTask -Id cve-exposure).Findings.Detail | Should -Match '\(Windows 11 Version 24H2 for x64-based Systems\)'
        }
        It 'works on the Windows PowerShell 5.1 parser output (dictionaries)' {
            InModuleScope AutoByte -Parameters @{ Doc = (ConvertTo-Dict ($script:CvrfJson | ConvertFrom-Json)) } {
                param($Doc)
                $r = @(Get-AbCveExposure -Documents @(, $Doc) -ProductNames 'Windows 10 Version 22H2 for x64-based Systems' -CurrentBuild ([version]'10.0.19045.5000'))
                $r.Cve | Sort-Object | Should -Be @('CVE-2026-1001', 'CVE-2026-1004')
            }
        }
        It 'fails soft when offline and can be turned off' {
            Mock -ModuleName AutoByte Get-AbCvrfDocument { }
            (Invoke-AbTask -Id cve-exposure).Findings.Code | Should -Be 'cve.data-unavailable'
            $env:AB_CveCheck = '0'
            try { (Invoke-AbTask -Id cve-exposure).Findings | Should -BeNullOrEmpty } finally { Remove-Item Env:AB_CveCheck }
        }
        It 'parses large JSON and reads fields from either shape' {
            InModuleScope AutoByte {
                $o = ConvertFrom-AbJson -Json '{"value":[{"ID":"2026-Nov"}]}'
                Get-AbField (@(Get-AbField $o 'value')[0]) 'ID' | Should -Be '2026-Nov'
                $d = New-Object 'System.Collections.Generic.Dictionary[string,object]'; $d['a'] = 1
                Get-AbField $d 'a' | Should -Be 1
                Get-AbField $d 'missing' | Should -BeNullOrEmpty
            }
        }
    }

    Context 'CVE feed download' {
        It 'reads the latest monthly documents from the feed index' {
            InModuleScope AutoByte {
                Mock Get-AbWebText { '{"value":[{"ID":"2026-Sep","InitialReleaseDate":"2026-09-09T07:00:00Z","CvrfUrl":"https://x/cvrf/2026-Sep"},{"ID":"2026-Nov","InitialReleaseDate":"2026-11-11T08:00:00Z","CvrfUrl":"https://x/cvrf/2026-Nov"},{"ID":"2026-Oct","InitialReleaseDate":"2026-10-14T07:00:00Z","CvrfUrl":"https://x/cvrf/2026-Oct"}]}' } -ParameterFilter { $Uri -like '*/updates' }
                Mock Get-AbWebText { '{"DocumentTitle":{"Value":"t"}}' } -ParameterFilter { $Uri -like '*/cvrf/*' }
                @(Get-AbCvrfDocument -BaseUrl 'https://x' -Months 2).Count | Should -Be 2
                Should -Invoke Get-AbWebText -ParameterFilter { $Uri -eq 'https://x/cvrf/2026-Nov' } -Times 1 -Exactly
                Should -Invoke Get-AbWebText -ParameterFilter { $Uri -eq 'https://x/cvrf/2026-Sep' } -Times 0 -Exactly
            }
        }
        It 'caches downloads and falls back to a stale cache when offline' {
            $old = $env:ProgramData
            $env:ProgramData = Join-Path ([IO.Path]::GetTempPath()) "ab-cache-$PID"
            try {
                InModuleScope AutoByte {
                    Mock Invoke-WebRequest { [pscustomobject]@{ Content = '{"v":1}' } }
                    Get-AbWebText -Uri 'https://x/a' -CacheName 't.json' | Should -Be '{"v":1}'
                    Get-AbWebText -Uri 'https://x/a' -CacheName 't.json' | Should -Be '{"v":1}'
                    Should -Invoke Invoke-WebRequest -Times 1 -Exactly -Because 'the second call is served from cache'
                    (Get-Item (Get-AbCachePath -Name 't.json')).LastWriteTime = (Get-Date).AddDays(-3)
                    Mock Invoke-WebRequest { throw 'offline' }
                    Get-AbWebText -Uri 'https://x/a' -CacheName 't.json' | Should -Be '{"v":1}'
                    Get-AbWebText -Uri 'https://x/b' -CacheName 'none.json' | Should -BeNullOrEmpty
                }
            }
            finally { Remove-Item -Recurse -Force $env:ProgramData -ErrorAction SilentlyContinue; $env:ProgramData = $old }
        }
    }

    Context 'app-updates and update-apps' {
        It 'parses winget output by column position' {
            InModuleScope AutoByte {
                $text = "   - `r   \ `rName                 Id                    Version   Available Source`n" +
                "------------------------------------------------------------------------`n" +
                "Google Chrome        Google.Chrome         130.0.1   131.0.2   winget`n" +
                "7-Zip 24.07 (x64)    7zip.7zip             24.07     24.08     winget`n" +
                "2 upgrades available."
                $a = @(ConvertFrom-AbWingetTable -Text $text)
                $a.Id | Should -Be @('Google.Chrome', '7zip.7zip')
                $a[1].Name | Should -Be '7-Zip 24.07 (x64)'
                $a[0].Available | Should -Be '131.0.2'
            }
        }
        It 'flags security-sensitive apps and handles a missing winget' {
            Mock -ModuleName AutoByte Get-AbAppUpgrade { [pscustomobject]@{ Name = 'Google Chrome'; Id = 'Google.Chrome'; Version = '1'; Available = '2' }; [pscustomobject]@{ Name = 'Spotify'; Id = 'Spotify.Spotify'; Version = '1'; Available = '2' } }
            $f = (Invoke-AbTask -Id app-updates).Findings
            $f.Severity | Should -Be 'Warning'
            $f.Title | Should -Match '2 app update\(s\) available, 1 security-sensitive'
            Mock -ModuleName AutoByte Get-AbAppUpgrade { $null }
            (Invoke-AbTask -Id app-updates).Findings.Code | Should -Be 'apps.winget-missing'
        }
        It 'upgrades each selected app by exact id, honouring -Exclude' {
            Mock -ModuleName AutoByte Get-AbWingetPath { 'C:\winget.exe' }
            $script:upgraded = @()
            Mock -ModuleName AutoByte Get-AbAppUpgrade { @([pscustomobject]@{ Name = 'A'; Id = 'Vendor.A'; Version = '1'; Available = '2' }, [pscustomobject]@{ Name = 'B'; Id = 'Vendor.B'; Version = '1'; Available = '2' }) | Where-Object { $script:upgraded -notcontains $_.Id } }
            Mock -ModuleName AutoByte Invoke-AbNative { $script:upgraded += $ArgumentList[2]; [pscustomobject]@{ ExitCode = 0; Output = '' } }
            (Invoke-AbTask -Id update-apps -Parameters @{ Exclude = 'Vendor.B' } -Confirm:$false).Status | Should -Be 'Fixed'
            Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains 'Vendor.A' -and $ArgumentList -contains '--exact' } -Times 1 -Exactly
            Should -Invoke -ModuleName AutoByte Invoke-AbNative -ParameterFilter { $ArgumentList -contains 'Vendor.B' } -Times 0 -Exactly
        }
    }

    Context 'repair-disk-errors' {
        It 'does nothing on a clean volume' {
            Mock -ModuleName AutoByte Get-AbVolumeScan { 'NoErrorsFound' }
            (Invoke-AbTask -Id repair-disk-errors -Confirm:$false).Status | Should -Be 'NotNeeded'
        }
        It 'applies an online spot fix when that is enough' {
            $script:fixed = $false
            Mock -ModuleName AutoByte Get-AbVolumeScan { if ($script:fixed) { 'NoErrorsFound' } else { 'SpotFixNeeded' } }
            Mock -ModuleName AutoByte Repair-AbVolume { $script:fixed = $true; 'NoErrorsFound' }
            $r = Invoke-AbTask -Id repair-disk-errors -Confirm:$false
            $r.Status | Should -Be 'Fixed'; $r.RebootRequired | Should -BeFalse
            Should -Invoke -ModuleName AutoByte Repair-AbVolume -ParameterFilter { $Mode -eq 'OfflineScanAndFix' } -Times 0 -Exactly
        }
        It 'schedules a full repair at restart for serious corruption' {
            Mock -ModuleName AutoByte Get-AbVolumeScan { 'FullRepairNeeded' }
            Mock -ModuleName AutoByte Repair-AbVolume { 'Scheduled' }
            $r = Invoke-AbTask -Id repair-disk-errors -Confirm:$false
            $r.Status | Should -Be 'Fixed'; $r.RebootRequired | Should -BeTrue
            Should -Invoke -ModuleName AutoByte Repair-AbVolume -ParameterFilter { $Mode -eq 'OfflineScanAndFix' } -Times 1 -Exactly
        }
    }

    Context 'every finding points at a real fix' {
        It 'all FixTask ids used by event rules and stop-code maps exist' {
            InModuleScope AutoByte {
                $fixes = @(Get-AbTask | Where-Object Kind -EQ 'Fix' | ForEach-Object Id)
                $ids = @($script:AbEventRules | Where-Object { $_.ContainsKey('FixTask') } | ForEach-Object FixTask) +
                @($script:AbStopCodes.Values | ForEach-Object { $_[2] }) + @($script:AbUpdateErrors.Values | ForEach-Object { $_[1] }) +
                'update-apps', 'repair-system-files', 'start-core-services', 'clear-print-queue', 'reset-windows-update', 'set-power-plan', 'rescan-devices', 'install-security-updates'
                foreach ($id in $ids | Where-Object { $_ } | Sort-Object -Unique) { $fixes | Should -Contain $id }
            }
        }
    }
}
