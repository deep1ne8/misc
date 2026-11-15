<#
.SYNOPSIS
    Runs the full AutoByte CI pipeline on this machine. No cloud service, account or billing.

.DESCRIPTION
    Same checks as .github/workflows/ci.yml, run locally:
      1. Leaked-credential guard      (build/Test-LeakedSecrets.ps1)
      2. Manifest is current          (build/Build-ScriptManifest.ps1 -Check)
      3. Parse every .ps1             (current PowerShell, plus Windows PowerShell 5.1 when available)
      4. PSScriptAnalyzer             (PSScriptAnalyzerSettings.psd1; errors fail)
      5. Pester                       (tests/powershell)
      6. Python: ruff, mypy, pytest   (in a private venv: .venv-ci)

    Missing modules (PSScriptAnalyzer, Pester) are installed for the current user only.
    Run it before pushing, or install the git hooks with -InstallHooks so it runs on every push.

.PARAMETER Skip
    Steps to skip: Secrets, Manifest, Parse, Analyzer, Pester, Python.

.PARAMETER InstallHooks
    Point git at .githooks/ (pre-commit: leaked-credential guard; pre-push: this script). Then exit.

.EXAMPLE
    pwsh ./build/Invoke-LocalCI.ps1
.EXAMPLE
    pwsh ./build/Invoke-LocalCI.ps1 -Skip Python
.EXAMPLE
    pwsh ./build/Invoke-LocalCI.ps1 -InstallHooks

.NOTES
    Version: 1.0.0
    Exit codes: 0 all steps passed, 1 one or more steps failed.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Secrets', 'Manifest', 'Parse', 'Analyzer', 'Pester', 'Python')]
    [string[]]$Skip = @(),
    [switch]$InstallHooks
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $Root

if ($InstallHooks) {
    & git config core.hooksPath .githooks
    if ($LASTEXITCODE -ne 0) { Write-Error 'git config failed.'; exit 1 }
    Write-Output 'Git hooks enabled: pre-commit runs the leaked-credential guard, pre-push runs this pipeline.'
    exit 0
}

$results = [System.Collections.Generic.List[object]]::new()
$psExe = (Get-Process -Id $PID).Path

function Invoke-Step {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Body)
    if ($Skip -contains $Name) { $results.Add([pscustomobject]@{ Step = $Name; Result = 'SKIPPED'; Seconds = 0; Detail = '' }); return }
    Write-Host "`n=== $Name ===" -ForegroundColor Cyan
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $detail = & $Body
        $results.Add([pscustomobject]@{ Step = $Name; Result = 'PASS'; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1); Detail = "$detail" })
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        $results.Add([pscustomobject]@{ Step = $Name; Result = 'FAIL'; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1); Detail = $_.Exception.Message })
    }
}

function Invoke-Child {
    # Runs a script in a separate PowerShell process so its 'exit' cannot end this one.
    param([string]$Exe = $psExe, [Parameter(Mandatory)][string[]]$Arguments)
    & $Exe -NoProfile -ExecutionPolicy Bypass @Arguments | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "$(Split-Path -Leaf $Arguments[1]) exited with $LASTEXITCODE" }
}

function Initialize-Module {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][version]$Minimum)
    if (-not (Get-Module -ListAvailable -Name $Name | Where-Object Version -GE $Minimum)) {
        Write-Host "Installing $Name >= $Minimum for the current user..." -ForegroundColor Yellow
        Install-Module -Name $Name -MinimumVersion $Minimum -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
    }
    Import-Module -Name $Name -MinimumVersion $Minimum -Force
}

$psFiles = @(Get-ChildItem -Recurse -Force -Include *.ps1, *.psm1 -File |
        Where-Object { $_.FullName -notmatch '[\\/](\.git|\.venv-ci|node_modules)[\\/]' -and $_.Length -gt 0 })

Invoke-Step 'Secrets' { Invoke-Child -Arguments '-File', (Join-Path $PSScriptRoot 'Test-LeakedSecrets.ps1') }

Invoke-Step 'Manifest' { Invoke-Child -Arguments '-File', (Join-Path $PSScriptRoot 'Build-ScriptManifest.ps1'), '-Check' }

Invoke-Step 'Parse' {
    $bad = 0
    foreach ($f in $psFiles) {
        $e = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$e)
        foreach ($x in @($e)) { $bad++; Write-Host "$($f.FullName):$($x.Extent.StartLineNumber) $($x.Message)" -ForegroundColor Red }
    }
    $note = "PowerShell $($PSVersionTable.PSVersion)"
    $winPs = if ($env:SystemRoot) { Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' } else { '' }
    if ($winPs -and (Test-Path -LiteralPath $winPs) -and $PSVersionTable.PSVersion.Major -ge 6) {
        $list = Join-Path ([IO.Path]::GetTempPath()) "autobyte-ci-files-$PID.txt"
        $psFiles.FullName | Set-Content -LiteralPath $list -Encoding UTF8
        $code = '$n=0; foreach($p in Get-Content -LiteralPath $args[0]){ $e=$null; [void][System.Management.Automation.Language.Parser]::ParseFile($p,[ref]$null,[ref]$e); foreach($x in @($e)){ $n++; Write-Output ("{0}:{1} {2}" -f $p,$x.Extent.StartLineNumber,$x.Message) } }; exit $n'
        & $winPs -NoProfile -ExecutionPolicy Bypass -Command $code $list | Out-Host
        $bad += $LASTEXITCODE
        Remove-Item -LiteralPath $list -Force -ErrorAction SilentlyContinue
        $note += ' + Windows PowerShell 5.1'
    }
    if ($bad) { throw "$bad parse error(s)" }
    "$($psFiles.Count) files ($note)"
}

Invoke-Step 'Analyzer' {
    Initialize-Module -Name PSScriptAnalyzer -Minimum '1.21.0'
    $settings = Join-Path $Root 'PSScriptAnalyzerSettings.psd1'
    $r = @($psFiles + @(Get-ChildItem -Recurse -Force -Filter *.psd1 -File | Where-Object FullName -NotMatch '[\\/]\.git[\\/]') |
            ForEach-Object { Invoke-ScriptAnalyzer -Path $_.FullName -Settings $settings })
    $errors = @($r | Where-Object Severity -EQ 'Error')
    foreach ($x in $errors) { Write-Host "$($x.ScriptPath):$($x.Line) $($x.RuleName): $($x.Message)" -ForegroundColor Red }
    if ($errors.Count) { throw "$($errors.Count) PSScriptAnalyzer error(s)" }
    "0 errors, $(@($r | Where-Object Severity -EQ 'Warning').Count) warnings"
}

Invoke-Step 'Pester' {
    $pesterScript = @'
Import-Module Pester -MinimumVersion 5.5.0
$c = New-PesterConfiguration
$c.Run.Path = 'tests/powershell'
$c.Run.Exit = $true
$c.Output.Verbosity = 'Normal'
Invoke-Pester -Configuration $c
'@
    Initialize-Module -Name Pester -Minimum '5.5.0'
    & $psExe -NoProfile -ExecutionPolicy Bypass -Command $pesterScript | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Pester reported $LASTEXITCODE failing test(s)" }
}

Invoke-Step 'Python' {
    $py = @('py', 'python3', 'python') | ForEach-Object { Get-Command $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
    if (-not $py) { throw 'Python 3.11+ not found. Install it from python.org or skip with -Skip Python.' }
    $venv = Join-Path $Root '.venv-ci'
    if (-not (Test-Path -LiteralPath $venv)) {
        $pre = @(); if ($py.Name -like 'py*.exe' -or $py.Name -eq 'py') { $pre = @('-3') }
        & $py.Source @pre -m venv $venv
        if ($LASTEXITCODE -ne 0) { throw 'Could not create .venv-ci' }
    }
    $vpy = @((Join-Path $venv 'Scripts/python.exe'), (Join-Path $venv 'bin/python')) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    & $vpy -m pip install --quiet --disable-pip-version-check -r AutoByte-Agent/requirements-dev.txt -r AutoByte-Agent/requirements-api.txt -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'pip install failed' }
    $checks = @(
        @{ Name = 'ruff (agent)'; Dir = $Root; Args = @('-m', 'ruff', 'check', 'AutoByte-Agent') },
        @{ Name = 'ruff format'; Dir = $Root; Args = @('-m', 'ruff', 'format', '--check', 'AutoByte-Agent') },
        @{ Name = 'ruff (GUI/tests)'; Dir = $Root; Args = @('-m', 'ruff', 'check', '--config', 'AutoByte-Agent/pyproject.toml', 'AutoByteGUI.py', 'tests/python') },
        @{ Name = 'mypy'; Dir = (Join-Path $Root 'AutoByte-Agent'); Args = @('-m', 'mypy') },
        @{ Name = 'pytest'; Dir = $Root; Args = @('-m', 'pytest', '-q', '-p', 'no:cacheprovider', 'AutoByte-Agent/tests', 'tests/python') }
    )
    foreach ($c in $checks) {
        Push-Location -LiteralPath $c.Dir
        try { & $vpy @($c.Args) | Out-Host; $code = $LASTEXITCODE } finally { Pop-Location }
        if ($code -ne 0) { throw "$($c.Name) failed ($code)" }
    }
    'ruff, mypy, pytest'
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Format-Table Step, Result, Seconds, Detail -AutoSize | Out-String -Width 200 | Write-Host
if ($results.Result -contains 'FAIL') { Write-Host 'LOCAL CI FAILED' -ForegroundColor Red; exit 1 }
Write-Host 'LOCAL CI PASSED' -ForegroundColor Green
exit 0
