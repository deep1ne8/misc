<#
.SYNOPSIS
    Datto RMM component for AutoByte v3: install/update, then scan or fix, with UDF and monitor output.

.DESCRIPTION
    Create a Datto RMM component (PowerShell) with this script. Optionally attach
    install/Install-AutoByte.ps1 to the component so endpoints do not need GitHub for the installer.

    Component variables (Datto passes them as environment variables):
      abAction   scan | fix | report | legacy          (default: scan)
      abTask     fix task id(s), comma-separated, or 'suggested'   (for abAction=fix)
      abScript   v2 script name                        (for abAction=legacy)
      abWhatIf   true | false                          (default: false) preview fixes only
      abFailOn   None | Warning | Critical             (default: None) make the job/monitor alert
      abUdf      1-30                                   (optional) write a one-line summary to that UDF
      abVersion  AutoByte version to install            (default: 3.1.0)
    Thresholds can be tuned with AB_* variables, e.g. AB_DiskWarnPercent=20 (see docs).

    Output: a human summary, the JSON result, and Datto monitor markers
    (<-Start Result-> / <-End Result-> and <-Start Diagnostic-> / <-End Diagnostic->).

.NOTES
    Version: 1.0.0
    Exit codes: 0 ok; 1 failure or findings matched abFailOn; 3010 restart required.
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Datto's agent may start 32-bit PowerShell; system tasks need the 64-bit host.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $native = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    & $native -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    exit $LASTEXITCODE
}

function Get-Var([string]$Name, [string]$Default = '') {
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { $Default } else { $v.Trim() }
}

$action = (Get-Var 'abAction' 'scan').ToLowerInvariant()
$version = Get-Var 'abVersion' '3.1.0'
$whatIf = (Get-Var 'abWhatIf' 'false') -match '^(true|1|yes)$'
$failOn = Get-Var 'abFailOn' 'None'
$udf = Get-Var 'abUdf' ''
if ($version -notmatch '^\d+\.\d+\.\d+$') { Write-Output "Invalid abVersion '$version'."; exit 1 }
if ($failOn -notin 'None', 'Warning', 'Critical') { Write-Output "Invalid abFailOn '$failOn'."; exit 1 }
if ($udf -and ($udf -notmatch '^\d+$' -or [int]$udf -lt 1 -or [int]$udf -gt 30)) { Write-Output "Invalid abUdf '$udf' (1-30)."; exit 1 }

# ---- install / update (hash-verified) -------------------------------------------------------
$installed = Get-Module -ListAvailable -Name AutoByte | Where-Object { $_.Version -eq [version]$version }
if (-not $installed) {
    $installer = Join-Path $PSScriptRoot 'Install-AutoByte.ps1'
    if (-not (Test-Path -LiteralPath $installer)) {
        $installer = Join-Path $env:TEMP "Install-AutoByte-$version.ps1"
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        (New-Object Net.WebClient).DownloadFile("https://raw.githubusercontent.com/deep1ne8/misc/v$version/install/Install-AutoByte.ps1", $installer)
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer -Version $version | Write-Output
    if ($LASTEXITCODE -ne 0) { Write-Output 'AutoByte install failed.'; exit 1 }
}
Import-Module AutoByte -RequiredVersion $version -Force

# ---- run ------------------------------------------------------------------------------------
$results = @(); $scan = $null
switch ($action) {
    { $_ -in 'scan', 'report' } {
        $scan = Invoke-AbScan
        $results = $scan.Results
        if ($action -eq 'report') { $file = Export-AbReport -Scan $scan; Write-Output "Report: $($file.FullName)" }
    }
    'fix' {
        $task = Get-Var 'abTask' ''
        if (-not $task) { Write-Output 'abAction=fix needs abTask (task id(s) or "suggested").'; exit 1 }
        $ids = if ($task -eq 'suggested') { (Invoke-AbScan).SuggestedFixes } else { $task -split '[,; ]+' | Where-Object { $_ } }
        foreach ($id in $ids) { $results += Invoke-AbTask -Id $id -WhatIf:$whatIf -Confirm:$false }
    }
    'legacy' {
        $name = Get-Var 'abScript' ''
        if (-not $name) { Write-Output 'abAction=legacy needs abScript.'; exit 1 }
        exit (Invoke-AbLegacyScript -Name $name -WhatIf:$whatIf)
    }
    default { Write-Output "Unknown abAction '$action' (scan, fix, report, legacy)."; exit 1 }
}

# ---- report back to Datto ---------------------------------------------------------------------
$findings = @($results | ForEach-Object { $_.Findings } | Where-Object { $_ })
$crit = @($findings | Where-Object Severity -EQ 'Critical').Count
$warn = @($findings | Where-Object Severity -EQ 'Warning').Count
$fixed = @($results | Where-Object Status -EQ 'Fixed').Count
$partial = @($results | Where-Object { $_.Status -eq 'Partial' -and $_.Kind -eq 'Fix' }).Count
$failed = @($results | Where-Object { $_.Status -eq 'Failed' -and $_.Kind -eq 'Fix' }).Count
$code = Get-AbExitCode -Results $results -FailOn $failOn
$status = if ($action -eq 'fix') { "AB $version fix: $fixed fixed, $partial partly, $failed failed" } else { "AB $version scan: $crit critical, $warn warning" }
$summary = "$status | $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

Write-Output $summary
foreach ($f in $findings | Where-Object Severity -NE 'Info') { Write-Output ("  [{0}] {1}{2}" -f $f.Severity, $f.Title, $(if ($f.FixTask) { "  (fix: $($f.FixTask))" } else { '' })) }
foreach ($r in $results | Where-Object Kind -EQ 'Fix') { Write-Output ("  {0}: {1} {2}" -f $r.TaskId, $r.Status, (($r.Messages + $r.Error) -join ' ')) }

if ($udf) {
    try { Set-ItemProperty -Path 'HKLM:\SOFTWARE\CentraStage' -Name "Custom$udf" -Value ($summary.Substring(0, [Math]::Min(255, $summary.Length))) -ErrorAction Stop }
    catch { Write-Output "Could not write UDF ${udf}: $($_.Exception.Message)" }
}

Write-Output '<-Start Diagnostic->'
if ($scan) { $scan | ConvertTo-AbJson -Compress | Write-Output } else { $results | ConvertTo-AbJson -Compress | Write-Output }
Write-Output '<-End Diagnostic->'
Write-Output '<-Start Result->'
Write-Output "STATUS=$status"
Write-Output '<-End Result->'
exit $code
