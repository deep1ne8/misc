#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }
# Engine contract tests: registry, Invoke-AbTask state machine, exit codes, config, report safety.

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
    $env:AUTOBYTE_LOG_DIR = Join-Path ([IO.Path]::GetTempPath()) "ab-test-logs-$PID"
    Import-Module (Join-Path $Root 'src/AutoByte/AutoByte.psd1') -Force
}
AfterAll { Remove-Module AutoByte -ErrorAction SilentlyContinue }

Describe 'Task registry' {
    It 'loads at least 14 checks and 8 fixes' {
        @(Get-AbTask -Kind Check).Count | Should -BeGreaterOrEqual 14
        @(Get-AbTask -Kind Fix).Count | Should -BeGreaterOrEqual 8
    }
    It 'every fix is marked mutating and every check is not' {
        Get-AbTask -Kind Fix | ForEach-Object { $_.Mutating | Should -BeTrue -Because $_.Id }
        Get-AbTask -Kind Check | ForEach-Object { $_.Mutating | Should -BeFalse -Because $_.Id }
    }
    It 'every finding FixTask refers to a real fix task' {
        $fixIds = @(Get-AbTask -Kind Fix).Id
        $src = Get-ChildItem (Join-Path $Root 'src/AutoByte/Tasks') -Recurse -Filter *.task.ps1 | Get-Content -Raw
        $refs = [regex]::Matches(($src -join "`n"), "-FixTask '([a-z0-9-]+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $refs | Where-Object { $_ -notin $fixIds } | Should -BeNullOrEmpty
    }
    It 'rejects malformed task definitions' {
        InModuleScope AutoByte {
            { Register-AbTask -Definition @{ Id = 'Bad_Id'; Name = 'x'; Kind = 'Check'; Category = 'x'; Description = 'x'; Test = {} } } | Should -Throw '*kebab*'
            { Register-AbTask -Definition @{ Id = 'no-repair'; Name = 'x'; Kind = 'Fix'; Category = 'x'; Description = 'x'; Test = {} } } | Should -Throw '*Repair*'
            { Register-AbTask -Definition @{ Id = 'disk-space'; Name = 'x'; Kind = 'Check'; Category = 'x'; Description = 'x'; Test = {} } } | Should -Throw '*Duplicate*'
        }
    }
}

Describe 'Invoke-AbTask state machine' {
    BeforeAll {
        InModuleScope AutoByte {
            $script:calls = @{ Repair = 0 }
            $script:state = @{ Needed = $true; RepairFixes = $true; Reboot = $false }
            Register-AbTask -Source 'test' -Definition @{
                Id = 'zz-test-fix'; Name = 'Test fix'; Kind = 'Fix'; Category = 'Test'; Description = 'test'; EndUserSafe = $false
                Parameters = @{ Level = 1 }
                Test = { param($Ctx, $P) @{ Needed = $script:state.Needed; Reason = 'state' } }
                Repair = { param($Ctx, $P) $script:calls.Repair++; if ($script:state.RepairFixes) { $script:state.Needed = $false }; 'stray output'; @{ RebootRequired = $script:state.Reboot } }
            }
            Register-AbTask -Source 'test' -Definition @{
                Id = 'zz-admin-check'; Name = 'Admin check'; Kind = 'Check'; Category = 'Test'; Description = 'test'; RequiresAdmin = $true
                Test = { param($Ctx, $P) New-AbFinding -Code 'x' -Severity Warning -Title 'x' }
            }
        }
    }
    BeforeEach {
        InModuleScope AutoByte { $script:calls.Repair = 0; $script:state.Needed = $true; $script:state.RepairFixes = $true; $script:state.Reboot = $false }
        Mock -ModuleName AutoByte Test-AbIsAdmin { $true }
    }
    AfterAll { InModuleScope AutoByte { $script:AbTasks.Remove('zz-test-fix'); $script:AbTasks.Remove('zz-admin-check') } }

    It 'does nothing when the fix is not needed (idempotent)' {
        InModuleScope AutoByte { $script:state.Needed = $false }
        (Invoke-AbTask -Id zz-test-fix -Confirm:$false).Status | Should -Be 'NotNeeded'
        InModuleScope AutoByte { $script:calls.Repair } | Should -Be 0
    }
    It 'repairs and verifies, ignoring stray pipeline output' {
        $r = Invoke-AbTask -Id zz-test-fix -Confirm:$false
        $r.Status | Should -Be 'Fixed'
        $r.Changed | Should -BeTrue
        InModuleScope AutoByte { $script:calls.Repair } | Should -Be 1
    }
    It 'honours -WhatIf and changes nothing' {
        (Invoke-AbTask -Id zz-test-fix -WhatIf).Status | Should -Be 'WhatIf'
        InModuleScope AutoByte { $script:calls.Repair } | Should -Be 0
    }
    It 'reports Failed when verification shows the problem remains' {
        InModuleScope AutoByte { $script:state.RepairFixes = $false }
        (Invoke-AbTask -Id zz-test-fix -Confirm:$false).Status | Should -Be 'Failed'
    }
    It 'treats a reboot-pending repair as Fixed with RebootRequired (exit 3010)' {
        InModuleScope AutoByte { $script:state.RepairFixes = $false; $script:state.Reboot = $true }
        $r = Invoke-AbTask -Id zz-test-fix -Confirm:$false
        $r.Status | Should -Be 'Fixed'
        $r.RebootRequired | Should -BeTrue
        InModuleScope AutoByte -Parameters @{ R = $r } { param($R) Get-AbExitCode -Results @($R) } | Should -Be 3010
    }
    It 'rejects unknown parameters' {
        $r = Invoke-AbTask -Id zz-test-fix -Parameters @{ Nope = 1 } -Confirm:$false
        $r.Status | Should -Be 'Failed'
        $r.Error | Should -BeLike '*Unknown parameter*'
    }
    It 'refuses admin-only fixes when not elevated, and skips admin-only checks' {
        Mock -ModuleName AutoByte Test-AbIsAdmin { $false }
        Mock -ModuleName AutoByte Get-AbService { }
        (Invoke-AbTask -Id clear-print-queue -Confirm:$false).Status | Should -Be 'Failed'
        (Invoke-AbTask -Id zz-admin-check).Status | Should -Be 'Skipped'
    }
    It 'throws for unknown task ids' {
        { Invoke-AbTask -Id does-not-exist } | Should -Throw '*Unknown task*'
    }
    It 'end-user mode skips fixes that need a technician' {
        $r = Invoke-AbFix -Id zz-test-fix -EndUserSafeOnly -Confirm:$false
        $r.Status | Should -Be 'Skipped'
        InModuleScope AutoByte { $script:calls.Repair } | Should -Be 0
    }
}

Describe 'Exit codes' {
    It 'maps results to 0 / 1 / 3010 and honours -FailOn' {
        InModuleScope AutoByte {
            $ok = [pscustomobject]@{ Kind = 'Check'; Status = 'Healthy'; RebootRequired = $false; Findings = @() }
            $warn = [pscustomobject]@{ Kind = 'Check'; Status = 'Issues'; RebootRequired = $false; Findings = @((New-AbFinding -Code a -Severity Warning -Title a)) }
            $checkErr = [pscustomobject]@{ Kind = 'Check'; Status = 'Failed'; RebootRequired = $false; Findings = @() }
            $fixErr = [pscustomobject]@{ Kind = 'Fix'; Status = 'Failed'; RebootRequired = $false; Findings = @() }
            $reboot = [pscustomobject]@{ Kind = 'Fix'; Status = 'Fixed'; RebootRequired = $true; Findings = @() }
            Get-AbExitCode -Results @($ok, $warn) | Should -Be 0
            Get-AbExitCode -Results @($checkErr) | Should -Be 0
            Get-AbExitCode -Results @($fixErr, $reboot) | Should -Be 1
            Get-AbExitCode -Results @($reboot) | Should -Be 3010
            Get-AbExitCode -Results @($warn) -FailOn Warning | Should -Be 1
            Get-AbExitCode -Results @($warn) -FailOn Critical | Should -Be 0
        }
    }
}

Describe 'Configuration' {
    It 'AB_* environment variables override defaults (Datto variables arrive this way)' {
        $env:AB_DiskWarnPercent = '30'
        try { InModuleScope AutoByte { (Get-AbConfig).DiskWarnPercent } | Should -Be 30 }
        finally { Remove-Item Env:AB_DiskWarnPercent }
    }
    It 'ignores non-numeric overrides for numeric settings' {
        $env:AB_DiskWarnPercent = 'lots'
        try { InModuleScope AutoByte { (Get-AbConfig).DiskWarnPercent } | Should -Be 15 }
        finally { Remove-Item Env:AB_DiskWarnPercent }
    }
}

Describe 'HTML report' {
    It 'is self-contained and HTML-encodes untrusted text' {
        $scan = [pscustomobject]@{
            Computer = 'PC<1>'; Summary = [pscustomobject]@{ OS = 'Windows 11'; Version = '10.0'; Manufacturer = 'Dell'; Model = 'X'; RamGB = 16; LastBoot = ''; AutoByte = '3.0.0' }
            Results = @(); Critical = 1; Warning = 0; Info = 0
            Findings = @(InModuleScope AutoByte { New-AbFinding -Code 'x' -Severity Critical -Title '<script>alert(1)</script>' -Detail 'a & b' -FixTask 'clean-temp-files' })
        }
        $path = Join-Path $TestDrive 'r.html'
        $null = Export-AbReport -Scan $scan -Path $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Not -Match '<script'
        $html | Should -Match '&lt;script&gt;'
        $html | Should -Match 'PC&lt;1&gt;'
        $html | Should -Not -Match '(src|href)="https?://'
        [IO.File]::ReadAllBytes($path)[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
    }
}
