#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }
# Console views, partial results, winget outcome mapping, Windows Sandbox and remote help.

BeforeAll {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
    $env:AUTOBYTE_LOG_DIR = Join-Path ([IO.Path]::GetTempPath()) "ab-test-logs-$PID"
    Import-Module (Join-Path $root 'src/AutoByte/AutoByte.psd1') -Force
    $script:Root = $root
    function script:Res($Id, $Status, $Messages = @(), $Data = @{}, $Kind = 'Fix', $Ms = 100) {
        [pscustomobject]@{ TaskId = $Id; Kind = $Kind; Status = $Status; Changed = $false; RebootRequired = $false; Findings = @(); Messages = $Messages; Error = ''; Data = $Data; DurationMs = $Ms; StartedUtc = '' }
    }
    function script:Capture([scriptblock]$Block) { (& $Block 6>&1 | ForEach-Object { "$_" }) -join "`n" }
}
AfterAll { Remove-Module AutoByte -ErrorAction SilentlyContinue }

Describe 'Console views' {
    BeforeEach { $null = Initialize-AbConsole -Mode None -Ascii }

    It 'uses plain-language status labels and hides engine wording' {
        $out = Capture { Write-AbFixView -Results @((Res 'clean-temp-files' 'NotNeeded' @('No action needed. Only 0.3 MB of old temp files.')), (Res 'resync-time' 'Fixed' @('Fix needed: clock is 90 s off.', 'Fixed and verified.'))) }
        $out | Should -Match 'Clean temporary files\s+Nothing to do'
        $out | Should -Match 'Only 0\.3 MB of old temp files\.'
        $out | Should -Not -Match 'No action needed'
        $out | Should -Not -Match 'Fixed and verified'
        $out | Should -Match 'Found: clock is 90 s off\.'
        $out | Should -Match '1 fixed\s+\|\s+1 nothing to do'
    }
    It 'shows per-app update results with reasons and next steps' {
        $data = @{
            Updated     = @([pscustomobject]@{ Id = 'a'; Name = '7-Zip'; From = '24'; To = '26.06'; Note = '' })
            Failed      = @([pscustomobject]@{ Id = 'n'; Name = 'Notepad++'; Code = '0x8A150101'; Reason = 'The app was open.'; Hint = 'Close it, then run the update again.' })
            StillListed = @([pscustomobject]@{ Id = 't'; Name = 'Windows Terminal' })
        }
        $out = Capture { Write-AbFixView -Results @(Res 'update-apps' 'Partial' @('Fix needed: 3 app(s)', 'Updated 7-Zip (24 -> 26.06)', 'Not updated: Notepad++ - x') $data -Ms 3602449) }
        $out | Should -Match 'Update applications\s+Partly fixed\s+1 h 0 min'
        $out | Should -Match '1 updated\s+\|\s+1 not updated\s+\|\s+1 still listed'
        $out | Should -Match '7-Zip 26\.06'
        $out | Should -Match 'x Notepad\+\+\s+The app was open\.'
        $out | Should -Match '-> Close it, then run the update again\.'
        $out | Should -Match 'Windows Terminal'
        $out | Should -Match 'Next\s+Follow the tips above, then run: abyte fix update-apps'
        $out | Should -Not -Match 'Not updated: Notepad'
    }
    It 'groups scan findings by severity and points end users at the right option' {
        $f = @(
            [pscustomobject]@{ Code = 'a'; Severity = 'Warning'; Title = 'Low disk space'; Detail = '9% free'; FixTask = 'clean-temp-files'; Data = @{} }
            [pscustomobject]@{ Code = 'b'; Severity = 'Critical'; Title = 'Missing security updates'; Detail = ''; FixTask = 'install-security-updates'; Data = @{} }
        )
        $scan = [pscustomobject]@{ Computer = 'PC1'; Summary = [pscustomobject]@{ OS = 'Windows 11 Pro' }; Results = @(Res 'disk-space' 'Issues' -Kind Check); Findings = $f; Critical = 1; Warning = 1; Info = 0; SuggestedFixes = @('clean-temp-files', 'install-security-updates') }
        $out = Capture { Write-AbScanView -Scan $scan -EndUser }
        $out.IndexOf('Needs attention now') | Should -BeLessThan $out.IndexOf('Worth fixing')
        $out | Should -Match 'Your IT support can fix this'
        $out | Should -Match 'AutoByte can fix this \(option 2\)'
        $out | Should -Match 'choose 2 to fix 1 safe item'
        $tech = Capture { Write-AbScanView -Scan $scan }
        $tech | Should -Match '-> abyte fix clean-temp-files'
        $tech | Should -Match 'Next\s+abyte fix --suggested'
    }
    It 'emits 24-bit pastel colours in TrueColor mode and none in plain mode' {
        $null = Initialize-AbConsole -Mode TrueColor -Ascii
        (Capture { Write-AbLine 'hi' Ok }) | Should -Match ([regex]::Escape("$([char]27)[38;2;168;230;207mhi"))
        $null = Initialize-AbConsole -Mode None -Ascii
        (Capture { Write-AbLine 'hi' Ok }) | Should -Be 'hi'
    }
    It 'draws a box whose sides line up' {
        $lines = (Capture { Write-AbBox -Lines @(@('AutoByte', 'Accent+', '  PC check-up', 'Soft'), @('PC1', 'Text')) }) -split "`n"
        $lines.Count | Should -Be 4
        @($lines | ForEach-Object Length | Select-Object -Unique).Count | Should -Be 1
    }
    It 'formats durations for people' {
        InModuleScope AutoByte {
            Format-AbDuration 120 | Should -Be '<1 s'
            Format-AbDuration 2500 | Should -Be '2.5 s'
            Format-AbDuration 75000 | Should -Be '1 min 15 s'
            Format-AbDuration 3602449 | Should -Be '1 h 0 min'
        }
    }
}

Describe 'Partial results and winget outcomes' {
    BeforeEach {
        Mock -ModuleName AutoByte Test-AbIsAdmin { $true }
        Mock -ModuleName AutoByte Get-AbWingetPath { 'C:\winget.exe' }
    }
    It 'maps winget exit codes to plain reasons and outcomes' {
        InModuleScope AutoByte {
            (Get-AbWingetOutcome -ExitCode -1978334975).Reason | Should -Be 'The app was open.'
            (Get-AbWingetOutcome -ExitCode -1978335107).Hint | Should -Match 'signed in as that user'
            (Get-AbWingetOutcome -ExitCode -1978335184).Code | Should -Be '0x8A150030'
            (Get-AbWingetOutcome -ExitCode -1978335189).Outcome | Should -Be 'Current'
            (Get-AbWingetOutcome -ExitCode 3010).Outcome | Should -Be 'Reboot'
            (Get-AbWingetOutcome -ExitCode 0).Outcome | Should -Be 'Ok'
            (Get-AbWingetOutcome -ExitCode 77).Outcome | Should -Be 'Fail'
        }
    }
    It 'reports Partial with per-app data, exit code 1, and treats "still listed" as a restart quirk' {
        $script:done = @()
        Mock -ModuleName AutoByte Get-AbAppUpgrade {
            @([pscustomobject]@{ Name = '7-Zip'; Id = '7zip.7zip'; Version = '24'; Available = '26' }
                [pscustomobject]@{ Name = 'Notepad++'; Id = 'Notepad++.Notepad++'; Version = '8'; Available = '9' }
                [pscustomobject]@{ Name = 'Windows Terminal'; Id = 'Microsoft.WindowsTerminal'; Version = '1.2'; Available = '1.25' }) |
                Where-Object { $script:done -notcontains $_.Id }
        }
        Mock -ModuleName AutoByte Invoke-AbNative {
            $id = $ArgumentList[2]
            if ($id -eq 'Notepad++.Notepad++') { return [pscustomobject]@{ ExitCode = -1978334975; Output = '' } }
            if ($id -eq '7zip.7zip') { $script:done += $id }
            [pscustomobject]@{ ExitCode = 0; Output = '' }    # Terminal "succeeds" but stays listed
        }
        $r = Invoke-AbTask -Id update-apps -Confirm:$false
        $r.Status | Should -Be 'Partial'
        @($r.Data.Updated).Count | Should -Be 2
        $r.Data.Failed[0].Reason | Should -Be 'The app was open.'
        $r.Data.StillListed.Name | Should -Be 'Windows Terminal'
        Get-AbExitCode -Results @($r) | Should -Be 1
    }
    It 'is Fixed when every app updates, and Failed (with data) when none do' {
        $script:done = @()
        Mock -ModuleName AutoByte Get-AbAppUpgrade { , @(@([pscustomobject]@{ Name = 'A'; Id = 'V.A'; Version = '1'; Available = '2' }) | Where-Object { $script:done -notcontains $_.Id }) }
        Mock -ModuleName AutoByte Invoke-AbNative { $script:done += $ArgumentList[2]; [pscustomobject]@{ ExitCode = 0; Output = '' } }
        (Invoke-AbTask -Id update-apps -Confirm:$false).Status | Should -Be 'Fixed'
        $script:done = @()
        Mock -ModuleName AutoByte Invoke-AbNative { [pscustomobject]@{ ExitCode = -1978335107; Output = '' } }
        $r = Invoke-AbTask -Id update-apps -Confirm:$false
        $r.Status | Should -Be 'Failed'
        $r.Changed | Should -BeFalse
        $r.Data.Failed[0].Code | Should -Be '0x8A15007D'
    }
}

Describe 'Windows Sandbox' {
    BeforeEach { $script:out = Join-Path ([IO.Path]::GetTempPath()) "ab-wsb-$([guid]::NewGuid().ToString('N'))" }
    AfterEach { Remove-Item -Recurse -Force $script:out -ErrorAction SilentlyContinue }

    It 'writes a config that maps AutoByte read-only and results read-write' {
        $cfg = New-AbSandboxConfig -Root $script:Root -OutputFolder $script:out -Arguments 'fix', 'reset-windows-update', '--yes' -NoNetwork -Confirm:$false
        [xml]$x = Get-Content -LiteralPath $cfg.FullName -Raw
        $x.Configuration.Networking | Should -Be 'Disable'
        $maps = @($x.Configuration.MappedFolders.MappedFolder)
        ($maps | Where-Object SandboxFolder -EQ 'C:\AutoByte').ReadOnly | Should -Be 'true'
        ($maps | Where-Object SandboxFolder -EQ 'C:\AutoByteResults').ReadOnly | Should -Be 'false'
        $x.Configuration.LogonCommand.Command | Should -Match 'C:\\AutoByteResults\\start\.ps1'
        $start = Get-Content -LiteralPath (Join-Path $script:out 'start.ps1') -Raw
        $start | Should -Match ([regex]::Escape("`$abArgs = @('fix', 'reset-windows-update', '--yes')"))
        $start | Should -Match 'AUTOBYTE_REPORT_DIR'
        $e = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($start, [ref]$null, [ref]$e); @($e).Count | Should -Be 0
    }
    It 'refuses arguments that could inject code' {
        foreach ($bad in "x'; Remove-Item C:\ -Recurse; '", 'a b', '$(calc)', 'x;y') {
            { New-AbSandboxConfig -Root $script:Root -OutputFolder $script:out -Arguments $bad -Confirm:$false } | Should -Throw '*Not allowed*'
        }
    }
    It 'explains why the sandbox is unavailable' {
        Mock -ModuleName AutoByte Test-AbPath { $false }
        Mock -ModuleName AutoByte Get-AbRegistryValue { 'Core' }
        (Get-AbSandboxStatus).Reason | Should -Match 'Pro, Enterprise or Education'
        Mock -ModuleName AutoByte Get-AbRegistryValue { 'Professional' }
        (Get-AbSandboxStatus).Reason | Should -Match 'abyte sandbox --enable'
        { Start-AbSandbox -Root $script:Root } | Should -Throw '*not turned on*'
    }
    It 'previews with -WhatIf without opening anything' {
        Mock -ModuleName AutoByte Get-AbSandboxStatus { [pscustomobject]@{ Available = $true; Running = $false; Exe = 'C:\WindowsSandbox.exe'; Reason = '' } }
        Mock -ModuleName AutoByte Start-Process { }
        Start-AbSandbox -Root $script:Root -WhatIf | Should -BeNullOrEmpty
        Should -Invoke -ModuleName AutoByte Start-Process -Times 0 -Exactly
    }
    It 'needs administrator to turn the feature on' {
        Mock -ModuleName AutoByte Test-AbIsAdmin { $false }
        { Enable-AbSandbox -Confirm:$false } | Should -Throw '*administrator*'
    }
}

Describe 'Remote help' {
    BeforeEach {
        Mock -ModuleName AutoByte Start-Process { }
        Mock -ModuleName AutoByte Get-AbService { }
    }
    AfterEach { Remove-Item Env:AB_RemoteSupportUrl -ErrorAction SilentlyContinue }
    It 'opens the MSP support page when configured' {
        $env:AB_RemoteSupportUrl = 'https://support.example.com/join'
        $r = Start-AbRemoteHelp -Confirm:$false
        $r.Method | Should -Be 'Url'
        Should -Invoke -ModuleName AutoByte Start-Process -ParameterFilter { $FilePath -eq 'https://support.example.com/join' } -Times 1 -Exactly
    }
    It 'refuses anything but https' {
        foreach ($u in 'http://support.example.com', 'javascript:alert(1)', 'file:///C:/Windows/System32/calc.exe') {
            $env:AB_RemoteSupportUrl = $u
            { Start-AbRemoteHelp -Confirm:$false } | Should -Throw '*https://*'
        }
        Should -Invoke -ModuleName AutoByte Start-Process -Times 0 -Exactly
    }
    It 'falls back to Quick Assist and mentions Datto Web Remote when the agent is present' {
        Mock -ModuleName AutoByte Get-AbService { [pscustomobject]@{ Name = 'CagService'; Status = 'Running' } }
        $r = Start-AbRemoteHelp -Confirm:$false
        $r.Method | Should -Be 'QuickAssist'
        $r.DattoAgent | Should -BeTrue
        Should -Invoke -ModuleName AutoByte Start-Process -ParameterFilter { $FilePath -eq 'ms-quick-assist:' } -Times 1 -Exactly
    }
}
