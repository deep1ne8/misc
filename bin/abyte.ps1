<#
.SYNOPSIS
    abyte - AutoByte command line: health checks, safe fixes and HTML reports.

.DESCRIPTION
    No arguments: a colourful, plain-language menu for end users (check my PC, fix what's safe,
    open the report, try fixes in Windows Sandbox, get remote help).

    Commands (technicians and RMM):
      abyte scan    [--json] [--report] [--open] [--fail-on warning|critical] [--category <name>]
      abyte check   <task-id>... [--param name=value] [--json] [--fail-on ...]   run single checks
      abyte fix     <task-id>... | --suggested   [--whatif] [--yes] [--param name=value] [--json]
      abyte run     <task-id>...                 runs checks (like check) or fixes (like fix)
      abyte list    [--kind check|fix] [--json]
      abyte report  [--open]                     scan and write the HTML report
      abyte sandbox [<abyte command...>] [--no-network] [--enable]
                                                 open AutoByte in Windows Sandbox (a throw-away
                                                 copy of Windows) to try fixes safely
      abyte remote  [--quick-assist]             get remote help (your IT's support page, or Quick Assist)
      abyte legacy  <ScriptName> [--whatif]      run a v2 library script (hash-verified)
      abyte version

    Display: --plain (no colours), --ascii (no Unicode symbols). NO_COLOR is honoured.
    Fixes change the PC, so they ask before running unless --yes is given. In a non-interactive
    session (RMM) --yes is required. Every fix supports --whatif.

.EXAMPLE
    abyte scan --report --open
.EXAMPLE
    abyte check cve-exposure event-logs
.EXAMPLE
    abyte fix clean-temp-files --whatif
.EXAMPLE
    abyte sandbox fix reset-windows-update --yes
.EXAMPLE
    abyte fix --suggested --yes --json

.NOTES
    Version: 3.1.0
    Exit codes: 0 ok; 1 a fix failed or only partly worked, or findings matched --fail-on;
    2 usage error; 3010 restart required.
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---- load the module (installed copy, else the one next to this script) -----------------------
$root = Split-Path -Parent $PSScriptRoot
$candidates = @(
    [IO.Path]::Combine($root, 'src', 'AutoByte', 'AutoByte.psd1')   # repository checkout
    [IO.Path]::Combine($root, 'Modules', 'AutoByte')                 # -InstallRoot layout
)
$local = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ($local) { Import-Module $local -Force } else { Import-Module AutoByte -ErrorAction Stop }

# ---- argument parsing (accepts --flag and -flag) ----------------------------------------------
$all = @($args | ForEach-Object { [string]$_ })
$style = @{ Mode = 'Auto'; Ascii = $false }
$all = @($all | Where-Object {
        if ($_ -in '--plain', '--no-color', '-plain') { $style.Mode = 'None'; $false }
        elseif ($_ -in '--ascii', '-ascii') { $style.Ascii = $true; $false }
        else { $true }
    })
$null = Initialize-AbConsole -Mode $style.Mode -Ascii:$style.Ascii

# 'sandbox' passes everything after it to abyte inside the sandbox, so it is parsed separately.
$sandbox = $null
if ($all.Count -and $all[0].ToLowerInvariant() -eq 'sandbox') {
    $sandbox = @{ Args = @(); NoNetwork = $false; Enable = $false }
    foreach ($a in @($all | Select-Object -Skip 1)) {
        if ($a -in '--no-network', '-no-network') { $sandbox.NoNetwork = $true }
        elseif ($a -in '--enable', '-enable') { $sandbox.Enable = $true }
        else { $sandbox.Args += $a }
    }
    $all = @('sandbox')
}

$positional = New-Object System.Collections.Generic.List[string]
$opt = @{ json = $false; report = $false; open = $false; whatif = $false; yes = $false; suggested = $false; quickassist = $false; failon = 'None'; category = ''; kind = ''; params = @{} }
for ($i = 0; $i -lt $all.Count; $i++) {
    $a = $all[$i]
    if ($a -notmatch '^--?([A-Za-z-]+)$') { $positional.Add($a); continue }
    $flag = $Matches[1].ToLowerInvariant() -replace '-', ''
    switch ($flag) {
        { $_ -in 'json', 'report', 'open', 'whatif', 'yes', 'suggested', 'quickassist' } { $opt[$flag] = $true; continue }
        'y' { $opt.yes = $true; continue }
        'failon' { $i++; $opt.failon = (Get-Culture).TextInfo.ToTitleCase(([string]$all[$i]).ToLowerInvariant()); continue }
        'category' { $i++; $opt.category = [string]$all[$i]; continue }
        'kind' { $i++; $opt.kind = [string]$all[$i]; continue }
        'param' {
            $i++
            $kv = [string]$all[$i]
            if ($kv -notmatch '^([A-Za-z]+)=(.*)$') { Write-Error "--param expects name=value, got '$kv'" -ErrorAction Continue; exit 2 }
            $pname = $Matches[1]; $v = $Matches[2]
            $n = 0
            if ($v -match '^(true|false)$') { $v = [bool]::Parse($v) } elseif ([int]::TryParse($v, [ref]$n)) { $v = $n }
            $opt.params[$pname] = $v
            continue
        }
        default { Write-Error "Unknown option '$a'. Run 'abyte help'." -ErrorAction Continue; exit 2 }
    }
}
$command = if ($positional.Count) { $positional[0].ToLowerInvariant() } else { '' }
$rest = @($positional | Select-Object -Skip 1)
$interactive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected

# ---- helpers ----------------------------------------------------------------------------------
function Invoke-Scan {
    $progress = { param($n, $total, $name) Write-Progress -Activity 'AutoByte health check' -Status "$n of ${total}: $name" -PercentComplete ([int](100 * ($n - 1) / $total)) }
    $p = @{ OnProgress = $progress }
    if ($opt.category) { $p.Category = $opt.category }
    $s = Invoke-AbScan @p
    Write-Progress -Activity 'AutoByte health check' -Completed
    $s
}
function Confirm-Fix([string[]]$Ids) {
    if ($opt.yes -or $opt.whatif) { return $true }
    if (-not $interactive) { Write-Error 'Fixes change the PC. In a non-interactive session pass --yes (or --whatif to preview).' -ErrorAction Continue; return $false }
    Write-Host ''
    Write-AbLine '  About to run' 'Accent+'
    foreach ($id in $Ids) {
        $t = Get-AbTask -Id $id
        Write-AbLine '   - ' Dim $t.Name 'Text+' $(if ($t.RebootMayBeRequired) { '  (may need a restart)' } else { '' }) Warn
        foreach ($l in @($t.Description)) { Write-AbLine '     ' Text $l Dim }
    }
    Write-Host ''
    (Read-Host '  Run these fixes? [y/N]') -match '^(y|yes)$'
}
function Invoke-FixSet([string[]]$Ids, [switch]$EndUser) {
    $results = @()
    $n = 0
    foreach ($id in $Ids) {
        $n++
        Write-Progress -Activity 'AutoByte fixes' -Status "$n of $($Ids.Count): $((Get-AbTask -Id $id).Name)" -PercentComplete ([int](100 * ($n - 1) / $Ids.Count))
        if ($opt.params.Count -and $Ids.Count -eq 1) {
            $results += Invoke-AbTask -Id $id -Parameters $opt.params -WhatIf:$opt.whatif -Confirm:$false
        }
        else { $results += Invoke-AbFix -Id $id -WhatIf:$opt.whatif -Confirm:$false -EndUserSafeOnly:$EndUser }
    }
    Write-Progress -Activity 'AutoByte fixes' -Completed
    $results
}
function Show-RemoteHelp {
    try {
        $r = Start-AbRemoteHelp -Method $(if ($opt.quickassist) { 'QuickAssist' } else { 'Auto' })
        $g = (Get-AbStatusView -Status Fixed).Glyph
        Write-Host ''
        Write-AbLine '  Remote help' 'Accent+'
        Write-AbLine '   ' Text "$g " Ok $r.Message Text
        Write-AbLine '   Tell your technician  ' Dim "Computer: $($r.Computer)" 'Key+' "   User: $($r.User)" Key
        if ($r.DattoAgent) { Write-AbLine '   ' Text 'Datto RMM is installed: your technician can also connect with Web Remote. Nothing else to install.' Info }
        0
    }
    catch {
        Write-Host ''
        Write-AbLine '  Remote help' 'Accent+'
        Write-AbLine '   ' Text $_.Exception.Message Bad
        1
    }
}
function Show-Sandbox([string[]]$SandboxArgs, [switch]$NoNetwork) {
    Write-Host ''
    Write-AbLine '  Windows Sandbox' 'Accent+' '  a throw-away copy of Windows; closing it discards every change' Dim
    $st = Get-AbSandboxStatus
    if (-not $st.Available -or $st.Running) { Write-AbLine '   ' Text $st.Reason Warn; return 1 }
    $r = Start-AbSandbox -Root $root -Arguments $SandboxArgs -NoNetwork:$NoNetwork -WhatIf:$opt.whatif
    if (-not $r) { Write-AbLine '   ' Text 'Preview only: the sandbox was not opened.' Info; return 0 }
    Write-AbLine '   Opening   ' Dim $(if ($SandboxArgs) { "abyte $($SandboxArgs -join ' ')" } else { 'the AutoByte menu' }) 'Key+' $(if ($NoNetwork) { '   (no network)' } else { '' }) Dim
    Write-AbLine '   Results   ' Dim $r.ResultsFolder Text
    Write-AbLine '   Note      ' Dim 'The sandbox is a clean Windows: app, driver and update checks show the sandbox, not this PC.' Dim
    0
}
function Show-Menu($Last) {
    $ver = (Get-Module AutoByte).Version.ToString()
    Clear-Host
    Write-Host ''
    Write-AbBox -Width 58 -Lines @(
        , @('AutoByte', 'Accent+', '  PC check-up', 'Soft', "              v$ver", 'Dim')
        , @([Environment]::MachineName, 'Text', "  $([Environment]::UserName)", 'Dim')
    )
    Write-Host ''
    $items = @(
        , @('1', 'Check my PC', 'looks for problems, changes nothing')
        , @('2', "Fix what's safe", 'temp files, network, printer, power plan')
        , @('3', 'Open the report', 'a page you can save or send to IT')
        , @('4', 'Try fixes in Windows Sandbox', 'a throw-away copy of Windows')
        , @('5', 'Get remote help', 'let your IT support connect')
        , @('Q', 'Quit', '')
    )
    foreach ($it in $items) { Write-AbLine '    ' Text " $($it[0]) " 'Key+' '  ' Text $it[1].PadRight(32) Text $it[2] Dim }
    Write-Host ''
    if ($Last) {
        $s = $Last
        $roleC = if ($s.Critical) { 'Bad+' } else { 'Dim' }; $roleW = if ($s.Warning) { 'Warn+' } else { 'Dim' }
        Write-AbLine '    Last check  ' Dim "$($s.Critical) critical" $roleC '  ' Text "$($s.Warning) warning(s)" $roleW '  ' Text "$($s.Info) note(s)" Info
    }
    else { Write-AbLine '    ' Text "Tip: start with 1. Nothing changes on your PC until you choose 2." Dim }
}

# ---- commands ---------------------------------------------------------------------------------
try {
    switch ($command) {
        '' {
            if (-not $interactive) { Write-Error 'No command given. Run: abyte help' -ErrorAction Continue; exit 2 }
            $lastScan = $null; $lastFix = @(); $lastReport = $null
            while ($true) {
                Show-Menu $lastScan
                $pause = $true
                $choice = (Read-Host "`n    Choose").Trim().ToUpperInvariant()
                switch ($choice) {
                    '1' {
                        $lastScan = Invoke-Scan
                        Write-AbScanView -Scan $lastScan -EndUser
                        $lastReport = Export-AbReport -Scan $lastScan
                        Write-AbLine '  Report  ' Dim $lastReport.FullName Text
                    }
                    '2' {
                        if (-not $lastScan) { $lastScan = Invoke-Scan }
                        $safe = @($lastScan.SuggestedFixes | Where-Object { (Get-AbTask -Id $_).EndUserSafe })
                        $tech = @($lastScan.SuggestedFixes | Where-Object { -not (Get-AbTask -Id $_).EndUserSafe })
                        if (-not $safe) { Write-Host ''; Write-AbLine '  ' Text 'Nothing that is safe to fix automatically.' Ok }
                        elseif (Confirm-Fix $safe) { $lastFix = @(Invoke-FixSet $safe -EndUser); Write-AbFixView -Results $lastFix }
                        if ($tech) {
                            Write-Host ''
                            Write-AbLine '  For your IT support' 'Soft+'
                            foreach ($t in $tech) { Write-AbLine '   - ' Dim (Get-AbTask -Id $t).Name Text }
                            Write-AbLine '  ' Text 'Choose 5 to let them connect.' Dim
                        }
                        $lastScan = Invoke-Scan
                        $lastReport = Export-AbReport -Scan $lastScan -FixResults $lastFix
                    }
                    '3' {
                        if (-not $lastReport) { $lastScan = Invoke-Scan; $lastReport = Export-AbReport -Scan $lastScan }
                        Start-Process -FilePath $lastReport.FullName
                        $pause = $false
                    }
                    '4' { $null = Show-Sandbox @() }
                    '5' { $null = Show-RemoteHelp }
                    'Q' { exit 0 }
                    default { Write-AbLine '    ' Text 'Please type 1, 2, 3, 4, 5 or Q.' Warn; Start-Sleep -Milliseconds 900; $pause = $false }
                }
                if ($pause) { Write-Host ''; $null = Read-Host '    Press Enter to go back to the menu' }
            }
        }
        { $_ -in 'help', '-h', '--help', '/?' } { Get-Help $PSCommandPath -Detailed | Out-Host; exit 0 }
        'version' { Write-Output (Get-Module AutoByte).Version.ToString(); exit 0 }
        'list' {
            $p = @{}; if ($opt.kind) { $p.Kind = (Get-Culture).TextInfo.ToTitleCase($opt.kind) }
            $tasks = @(Get-AbTask @p)
            if ($opt.json) { $tasks | ConvertTo-AbJson; exit 0 }
            foreach ($kind in 'Check', 'Fix') {
                $set = @($tasks | Where-Object Kind -EQ $kind | Sort-Object Category, Id)
                if (-not $set) { continue }
                Write-Host ''
                Write-AbRule -Title $(if ($kind -eq 'Check') { 'Checks (read-only)' } else { 'Fixes' })
                foreach ($cat in @($set | ForEach-Object Category | Select-Object -Unique)) {
                    Write-AbLine '  ' Text $cat 'Soft+'
                    foreach ($t in @($set | Where-Object Category -EQ $cat)) {
                        $tags = @(); if ($t.RequiresAdmin) { $tags += 'admin' }; if ($t.EndUserSafe) { $tags += 'end-user safe' }; if ($t.RebootMayBeRequired) { $tags += 'may restart' }
                        Write-AbLine '    ' Text $t.Id.PadRight(28) Key $t.Name Text $(if ($tags) { "   [$($tags -join ', ')]" } else { '' }) Dim
                    }
                }
            }
            exit 0
        }
        { $_ -in 'scan', 'report' } {
            $s = Invoke-Scan
            if ($command -eq 'report') { $opt.report = $true }
            $file = $null
            if ($opt.report -or $opt.open) { $file = Export-AbReport -Scan $s -Open:$opt.open }
            if ($opt.json) { $s | ConvertTo-AbJson }
            else { Write-AbScanView -Scan $s; if ($file) { Write-AbLine '  Report  ' Dim $file.FullName Text } }
            exit (Get-AbExitCode -Results $s.Results -FailOn $opt.failon)
        }
        { $_ -eq 'check' -or ($_ -eq 'run' -and $rest -and -not @($rest | Where-Object { (Get-AbTask -Id $_).Kind -ne 'Check' })) } {
            if (-not $rest) { Write-Error "Usage: abyte check <task-id>... Run 'abyte list --kind check'." -ErrorAction Continue; exit 2 }
            $ids = @($rest | Select-Object -Unique)
            foreach ($id in $ids) { if ((Get-AbTask -Id $id).Kind -ne 'Check') { Write-Error "'$id' is a fix, not a check. Use 'abyte fix $id'." -ErrorAction Continue; exit 2 } }
            $results = @()
            $n = 0
            foreach ($id in $ids) {
                $n++
                Write-Progress -Activity 'AutoByte checks' -Status "$n of $($ids.Count): $((Get-AbTask -Id $id).Name)" -PercentComplete ([int](100 * ($n - 1) / $ids.Count))
                $results += if ($opt.params.Count -and $ids.Count -eq 1) { Invoke-AbTask -Id $id -Parameters $opt.params } else { Invoke-AbTask -Id $id }
            }
            Write-Progress -Activity 'AutoByte checks' -Completed
            if ($opt.json) { $results | ConvertTo-AbJson }
            else {
                $f = @($results | ForEach-Object { $_.Findings } | Where-Object { $_ })
                Write-AbScanView -Scan ([pscustomobject]@{
                        Computer = [Environment]::MachineName; Summary = $null; Results = $results; Findings = $f
                        Critical = @($f | Where-Object Severity -EQ 'Critical').Count; Warning = @($f | Where-Object Severity -EQ 'Warning').Count; Info = @($f | Where-Object Severity -EQ 'Info').Count
                        SuggestedFixes = @($f | Where-Object { $_.FixTask -and $_.Severity -ne 'Info' } | ForEach-Object FixTask | Sort-Object -Unique)
                    })
            }
            exit (Get-AbExitCode -Results $results -FailOn $opt.failon)
        }
        { $_ -in 'fix', 'run' } {
            $ids = @($rest)
            if ($opt.suggested) { $ids += @((Invoke-Scan).SuggestedFixes) }
            $ids = @($ids | Where-Object { $_ } | Sort-Object -Unique)
            if (-not $ids) { Write-AbLine '  ' Text 'Nothing to fix.' Ok; exit 0 }
            foreach ($id in $ids) { if ((Get-AbTask -Id $id).Kind -ne 'Fix') { Write-Error "'$id' is not a fix task. Run 'abyte list --kind fix'." -ErrorAction Continue; exit 2 } }
            if (-not (Confirm-Fix $ids)) { exit 2 }
            $results = @(Invoke-FixSet $ids)
            if ($opt.json) { $results | ConvertTo-AbJson } else { Write-AbFixView -Results $results }
            exit (Get-AbExitCode -Results $results)
        }
        'sandbox' {
            if ($sandbox.Enable) {
                $restart = Enable-AbSandbox -WhatIf:$opt.whatif
                Write-AbLine '  ' Text $(if ($restart) { 'Windows Sandbox is turned on. Restart the PC, then run: abyte sandbox' } else { 'Windows Sandbox is ready. Run: abyte sandbox' }) Ok
                exit $(if ($restart) { 3010 } else { 0 })
            }
            exit (Show-Sandbox -SandboxArgs $sandbox.Args -NoNetwork:$sandbox.NoNetwork)
        }
        'remote' { exit (Show-RemoteHelp) }
        'legacy' {
            if (-not $rest) { Write-Error 'Usage: abyte legacy <ScriptName> [--whatif]' -ErrorAction Continue; exit 2 }
            exit (Invoke-AbLegacyScript -Name $rest[0] -WhatIf:$opt.whatif)
        }
        default { Write-Error "Unknown command '$command'. Run 'abyte help'." -ErrorAction Continue; exit 2 }
    }
}
catch {
    Write-Error $_.Exception.Message -ErrorAction Continue
    exit 1
}
