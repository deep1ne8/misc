<#
.SYNOPSIS
    abyte - AutoByte command line: health checks, safe fixes and HTML reports.

.DESCRIPTION
    No arguments: a plain-language menu for end users (check my PC, fix what's safe, open report).

    Commands (technicians and RMM):
      abyte scan   [--json] [--report] [--open] [--fail-on warning|critical] [--category <name>]
      abyte check  <task-id>... [--param name=value] [--json] [--fail-on ...]   run single checks
      abyte fix    <task-id>... | --suggested   [--whatif] [--yes] [--param name=value] [--json]
      abyte run    <task-id>...                   runs checks (like check) or fixes (like fix)
      abyte list   [--kind check|fix] [--json]
      abyte report [--open]                       scan and write the HTML report
      abyte legacy <ScriptName> [--whatif]        run a v2 library script (hash-verified)
      abyte version

    Fixes change the PC, so they ask before running unless --yes is given. In a non-interactive
    session (RMM) --yes is required. Every fix supports --whatif.

.EXAMPLE
    abyte scan --report --open
.EXAMPLE
    abyte check cve-exposure event-logs
.EXAMPLE
    abyte fix clean-temp-files --whatif
.EXAMPLE
    abyte fix --suggested --yes --json

.NOTES
    Version: 3.0.0
    Exit codes: 0 ok; 1 a fix failed, or findings matched --fail-on; 2 usage error; 3010 restart required.
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---- load the module (installed copy, else the one next to this script) -----------------------
$candidates = @(
    [IO.Path]::Combine((Split-Path -Parent $PSScriptRoot), 'src', 'AutoByte', 'AutoByte.psd1')   # repository checkout
    [IO.Path]::Combine((Split-Path -Parent $PSScriptRoot), 'Modules', 'AutoByte')                 # -InstallRoot layout
)
$local = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ($local) { Import-Module $local -Force } else { Import-Module AutoByte -ErrorAction Stop }

# ---- argument parsing (accepts --flag and -flag) ----------------------------------------------
$positional = New-Object System.Collections.Generic.List[string]
$opt = @{ json = $false; report = $false; open = $false; whatif = $false; yes = $false; suggested = $false; failon = 'None'; category = ''; kind = ''; params = @{} }
for ($i = 0; $i -lt $args.Count; $i++) {
    $a = [string]$args[$i]
    if ($a -notmatch '^--?([A-Za-z-]+)$') { $positional.Add($a); continue }
    $flag = $Matches[1].ToLowerInvariant() -replace '-', ''
    switch ($flag) {
        { $_ -in 'json', 'report', 'open', 'whatif', 'yes', 'suggested' } { $opt[$flag] = $true; continue }
        'y' { $opt.yes = $true; continue }
        'failon' { $i++; $opt.failon = (Get-Culture).TextInfo.ToTitleCase(([string]$args[$i]).ToLowerInvariant()); continue }
        'category' { $i++; $opt.category = [string]$args[$i]; continue }
        'kind' { $i++; $opt.kind = [string]$args[$i]; continue }
        'param' {
            $i++
            $kv = [string]$args[$i]
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

# ---- output helpers ---------------------------------------------------------------------------
$colors = @{ Critical = 'Red'; Warning = 'Yellow'; Info = 'Cyan'; Healthy = 'Green'; Fixed = 'Green'; NotNeeded = 'DarkGreen'; Failed = 'Red'; WhatIf = 'Cyan'; Skipped = 'DarkGray'; Issues = 'Yellow' }
function Write-Finding($f) {
    $c = $colors[$f.Severity]
    Write-Host ("  [{0,-8}] {1}" -f $f.Severity, $f.Title) -ForegroundColor $c
    if ($f.Detail) { Write-Host "             $($f.Detail)" -ForegroundColor DarkGray }
    if ($f.FixTask) { Write-Host "             fix: abyte fix $($f.FixTask)" -ForegroundColor DarkGray }
}
function Write-FixResult($r) {
    $name = (Get-AbTask -Id $r.TaskId).Name
    Write-Host ("  [{0,-9}] {1}" -f $r.Status, $name) -ForegroundColor $colors[$r.Status]
    foreach ($m in @($r.Messages)) { Write-Host "              $m" -ForegroundColor DarkGray }
    if ($r.Error) { Write-Host "              $($r.Error)" -ForegroundColor Red }
}
function Invoke-Scan {
    $progress = { param($n, $total, $name) Write-Progress -Activity 'AutoByte health check' -Status "$n/${total}: $name" -PercentComplete ([int](100 * ($n - 1) / $total)) }
    $p = @{ OnProgress = $progress }
    if ($opt.category) { $p.Category = $opt.category }
    $s = Invoke-AbScan @p
    Write-Progress -Activity 'AutoByte health check' -Completed
    $s
}
function Show-Scan($s) {
    Write-Host ''
    Write-Host "AutoByte health check - $($s.Computer)" -ForegroundColor White
    Write-Host ("  {0} critical, {1} warning(s), {2} note(s) from {3} checks" -f $s.Critical, $s.Warning, $s.Info, @($s.Results).Count)
    $order = @{ Critical = 0; Warning = 1; Info = 2 }
    foreach ($f in ($s.Findings | Sort-Object { $order[$_.Severity] })) { Write-Finding $f }
    if (-not $s.Findings) { Write-Host '  Everything looks good.' -ForegroundColor Green }
    $err = @($s.Results | Where-Object Status -EQ 'Failed')
    foreach ($r in $err) { Write-Host "  (check '$($r.TaskId)' could not run: $($r.Error))" -ForegroundColor DarkGray }
}
function Confirm-Fix([string[]]$Ids) {
    if ($opt.yes -or $opt.whatif) { return $true }
    if (-not $interactive) { Write-Error 'Fixes change the PC. In a non-interactive session pass --yes (or --whatif to preview).' -ErrorAction Continue; return $false }
    Write-Host ''
    foreach ($id in $Ids) { $t = Get-AbTask -Id $id; Write-Host "  - $($t.Name): $($t.Description)" }
    (Read-Host 'Run these fixes? [y/N]') -match '^(y|yes)$'
}
function Invoke-FixSet([string[]]$Ids, [switch]$EndUser) {
    $results = @()
    $n = 0
    foreach ($id in $Ids) {
        $n++
        Write-Progress -Activity 'AutoByte fixes' -Status "$n/$($Ids.Count): $((Get-AbTask -Id $id).Name)" -PercentComplete ([int](100 * ($n - 1) / $Ids.Count))
        $p = @{ Id = $id; WhatIf = $opt.whatif; Confirm = $false; EndUserSafeOnly = [bool]$EndUser }
        if ($opt.params.Count -and $Ids.Count -eq 1) {
            $results += Invoke-AbTask -Id $id -Parameters $opt.params -WhatIf:$opt.whatif -Confirm:$false
        }
        else { $results += Invoke-AbFix @p }
    }
    Write-Progress -Activity 'AutoByte fixes' -Completed
    $results
}

# ---- commands ---------------------------------------------------------------------------------
try {
    switch ($command) {
        '' {
            if (-not $interactive) { Write-Error 'No command given. Run: abyte help' -ErrorAction Continue; exit 2 }
            $lastScan = $null; $lastFix = @(); $lastReport = $null
            while ($true) {
                Write-Host ''
                Write-Host '  AutoByte - PC check-up' -ForegroundColor Cyan
                Write-Host '  1  Check my PC (makes no changes)'
                Write-Host '  2  Fix the problems that are safe to fix'
                Write-Host '  3  Open the report'
                Write-Host '  Q  Quit'
                $choice = (Read-Host '  Choose').Trim().ToUpperInvariant()
                switch ($choice) {
                    '1' {
                        $lastScan = Invoke-Scan; Show-Scan $lastScan
                        $lastReport = Export-AbReport -Scan $lastScan
                        Write-Host "`n  Report saved: $($lastReport.FullName)" -ForegroundColor DarkGray
                    }
                    '2' {
                        if (-not $lastScan) { $lastScan = Invoke-Scan }
                        $safe = @($lastScan.SuggestedFixes | Where-Object { (Get-AbTask -Id $_).EndUserSafe })
                        $tech = @($lastScan.SuggestedFixes | Where-Object { -not (Get-AbTask -Id $_).EndUserSafe })
                        if (-not $safe) { Write-Host '  Nothing to fix automatically.' -ForegroundColor Green }
                        elseif (Confirm-Fix $safe) { $lastFix = Invoke-FixSet $safe -EndUser; $lastFix | ForEach-Object { Write-FixResult $_ } }
                        if ($tech) { Write-Host "`n  These need your IT support: $((($tech | ForEach-Object { (Get-AbTask -Id $_).Name }) -join ', '))" -ForegroundColor Yellow }
                        $lastScan = Invoke-Scan
                        $lastReport = Export-AbReport -Scan $lastScan -FixResults $lastFix
                    }
                    '3' {
                        if (-not $lastReport) { $lastScan = Invoke-Scan; $lastReport = Export-AbReport -Scan $lastScan }
                        Start-Process -FilePath $lastReport.FullName
                    }
                    'Q' { exit 0 }
                    default { Write-Host '  Please type 1, 2, 3 or Q.' -ForegroundColor Yellow }
                }
            }
        }
        { $_ -in 'help', '-h', '--help', '/?' } { Get-Help $PSCommandPath -Detailed | Out-Host; exit 0 }
        'version' { Write-Output (Get-Module AutoByte).Version.ToString(); exit 0 }
        'list' {
            $p = @{}; if ($opt.kind) { $p.Kind = (Get-Culture).TextInfo.ToTitleCase($opt.kind) }
            $tasks = @(Get-AbTask @p)
            if ($opt.json) { $tasks | ConvertTo-AbJson; exit 0 }
            $tasks | Sort-Object Kind, Category, Id |
                Select-Object Kind, Category, Id, Name, @{ n = 'Admin'; e = { if ($_.RequiresAdmin) { 'yes' } else { '' } } } |
                Format-Table -AutoSize | Out-String -Width 160 | Write-Host
            exit 0
        }
        { $_ -in 'scan', 'report' } {
            $s = Invoke-Scan
            if ($command -eq 'report') { $opt.report = $true }
            $file = $null
            if ($opt.report -or $opt.open) { $file = Export-AbReport -Scan $s -Open:$opt.open }
            if ($opt.json) { $s | ConvertTo-AbJson }
            else { Show-Scan $s; if ($file) { Write-Host "`n  Report: $($file.FullName)" -ForegroundColor DarkGray } }
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
                Write-Progress -Activity 'AutoByte checks' -Status "$n/$($ids.Count): $((Get-AbTask -Id $id).Name)" -PercentComplete ([int](100 * ($n - 1) / $ids.Count))
                $results += if ($opt.params.Count -and $ids.Count -eq 1) { Invoke-AbTask -Id $id -Parameters $opt.params } else { Invoke-AbTask -Id $id }
            }
            Write-Progress -Activity 'AutoByte checks' -Completed
            if ($opt.json) { $results | ConvertTo-AbJson }
            else {
                foreach ($r in $results) {
                    Write-Host ''
                    Write-Host ("{0} [{1}]" -f (Get-AbTask -Id $r.TaskId).Name, $r.Status) -ForegroundColor $colors[$r.Status]
                    foreach ($m in @($r.Messages)) { Write-Host "  $m" -ForegroundColor DarkGray }
                    if ($r.Error) { Write-Host "  could not run: $($r.Error)" -ForegroundColor Red }
                    $order = @{ Critical = 0; Warning = 1; Info = 2 }
                    foreach ($f in @($r.Findings | Sort-Object { $order[$_.Severity] })) { Write-Finding $f }
                    if ($r.Status -eq 'Healthy' -and -not @($r.Findings)) { Write-Host '  No problems found.' -ForegroundColor Green }
                }
            }
            exit (Get-AbExitCode -Results $results -FailOn $opt.failon)
        }
        { $_ -in 'fix', 'run' } {
            $ids = @($rest)
            if ($opt.suggested) { $ids += @((Invoke-Scan).SuggestedFixes) }
            $ids = @($ids | Where-Object { $_ } | Sort-Object -Unique)
            if (-not $ids) { Write-Host 'Nothing to fix.' -ForegroundColor Green; exit 0 }
            foreach ($id in $ids) { if ((Get-AbTask -Id $id).Kind -ne 'Fix') { Write-Error "'$id' is not a fix task. Run 'abyte list --kind fix'." -ErrorAction Continue; exit 2 } }
            if (-not (Confirm-Fix $ids)) { exit 2 }
            $results = @(Invoke-FixSet $ids)
            if ($opt.json) { $results | ConvertTo-AbJson } else { Write-Host ''; $results | ForEach-Object { Write-FixResult $_ } }
            exit (Get-AbExitCode -Results $results)
        }
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
