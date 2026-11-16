# Core contract: findings, results, logging, configuration and the task registry.

$script:Severities = @('Info', 'Warning', 'Critical')

function New-AbFinding {
    <#
    .SYNOPSIS
        Creates a finding: one problem (or notable fact) detected by a check.
    .PARAMETER Code
        Stable machine-readable id, e.g. 'disk.low-space'. RMM monitors and the AI layer key on it.
    .PARAMETER FixTask
        Id of the fix task that resolves this finding, if any.
    #>
    param(
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][ValidateSet('Info', 'Warning', 'Critical')][string]$Severity,
        [Parameter(Mandatory)][string]$Title,
        [string]$Detail = '',
        [string]$FixTask = '',
        [hashtable]$Data = @{}
    )
    [pscustomobject]@{
        PSTypeName = 'AutoByte.Finding'
        Code       = $Code
        Severity   = $Severity
        Title      = $Title
        Detail     = $Detail
        FixTask    = $FixTask
        Data       = $Data
    }
}

function New-AbResult {
    param([Parameter(Mandatory)][string]$TaskId, [Parameter(Mandatory)][string]$Kind)
    [pscustomobject]@{
        PSTypeName     = 'AutoByte.Result'
        TaskId         = $TaskId
        Kind           = $Kind
        Status         = 'Pending'   # Healthy | Issues | NotNeeded | Fixed | Failed | Skipped | WhatIf
        Changed        = $false
        RebootRequired = $false
        Findings       = @()
        Messages       = @()
        Error          = ''
        DurationMs     = 0
        StartedUtc     = (Get-Date).ToUniversalTime().ToString('o')
    }
}

function Add-AbMessage {
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][string]$Message)
    $Result.Messages = @($Result.Messages) + $Message
    Write-AbLog -Message "[$($Result.TaskId)] $Message"
}

function Get-AbLogPath {
    $base = if ($env:ProgramData) { $env:ProgramData } else { [IO.Path]::GetTempPath() }
    if ($env:AUTOBYTE_LOG_DIR) { $dir = $env:AUTOBYTE_LOG_DIR } else { $dir = Join-Path $base 'AutoByte\Logs' }
    Join-Path $dir ('AutoByte_{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
}

function Write-AbLog {
    # Best-effort file log. Never throws; never write secrets here.
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    try {
        $path = Get-AbLogPath
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        [IO.File]::AppendAllText($path, $line + [Environment]::NewLine)
    }
    catch { Write-Verbose "Log write failed: $_" }
}

function Get-AbConfig {
    <#
    .SYNOPSIS
        Effective settings: built-in defaults, overridden by %ProgramData%\AutoByte\config.json,
        overridden by AB_* environment variables (Datto RMM site/component variables arrive this way).
        Never store secrets in config.json; pass them as masked RMM variables instead.
    #>
    $cfg = @{
        DiskWarnPercent    = 15
        DiskCriticalPercent = 5
        MemoryWarnPercent  = 10
        UptimeWarnDays     = 14
        EventErrorWarn     = 25
        EventLookbackDays  = 7
        CpuWarnPercent     = 85
        CommitWarnPercent  = 90
        DiskBusyWarnPercent = 90
        BootWarnSeconds    = 120
        StartupItemsWarn   = 15
        CveCheck           = 1
        CveMonths          = 3
        CveFeedUrl         = 'https://api.msrc.microsoft.com/cvrf/v3.0'
        KevFeedUrl         = 'https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json'
        DefenderSigMaxDays = 3
        TimeSkewWarnSeconds = 60
        TempFileMinAgeDays = 7
        ConnectivityHosts  = 'www.microsoft.com,login.microsoftonline.com'
    }
    $base = if ($env:ProgramData) { $env:ProgramData } else { [IO.Path]::GetTempPath() }
    $file = Join-Path $base 'AutoByte\config.json'
    if (Test-Path -LiteralPath $file) {
        try {
            $json = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) { if ($cfg.ContainsKey($p.Name)) { $cfg[$p.Name] = $p.Value } }
        }
        catch { Write-AbLog -Level WARN -Message "Ignoring unreadable config.json: $($_.Exception.Message)" }
    }
    foreach ($k in @($cfg.Keys)) {
        $v = [Environment]::GetEnvironmentVariable("AB_$k")
        if (-not [string]::IsNullOrWhiteSpace($v)) {
            if ($cfg[$k] -is [int]) { $n = 0; if ([int]::TryParse($v, [ref]$n)) { $cfg[$k] = $n } } else { $cfg[$k] = $v }
        }
    }
    $cfg
}

# ---- task registry ---------------------------------------------------------------------------
$script:AbTasks = [ordered]@{}

function Register-AbTask {
    param([Parameter(Mandatory)][hashtable]$Definition, [string]$Source = '')
    foreach ($k in 'Id', 'Name', 'Kind', 'Category', 'Description', 'Test') {
        if (-not $Definition.ContainsKey($k)) { throw "Task in '$Source' is missing '$k'." }
    }
    if ($Definition.Id -notmatch '^[a-z0-9]+(-[a-z0-9]+)*$') { throw "Task id '$($Definition.Id)' must be lowercase-kebab-case." }
    if ($Definition.Kind -notin 'Check', 'Fix') { throw "Task '$($Definition.Id)': Kind must be Check or Fix." }
    if ($Definition.Kind -eq 'Fix' -and -not $Definition.ContainsKey('Repair')) { throw "Fix task '$($Definition.Id)' has no Repair block." }
    if ($script:AbTasks.Contains($Definition.Id)) { throw "Duplicate task id '$($Definition.Id)'." }
    $defaults = @{ Mutating = ($Definition.Kind -eq 'Fix'); RequiresAdmin = $false; EndUserSafe = $false; RebootMayBeRequired = $false; Parameters = @{}; Tags = @() }
    foreach ($k in $defaults.Keys) { if (-not $Definition.ContainsKey($k)) { $Definition[$k] = $defaults[$k] } }
    $Definition.Source = $Source
    $script:AbTasks[$Definition.Id] = $Definition
}

function Get-AbTaskDefinition {
    param([Parameter(Mandatory)][string]$Id)
    if (-not $script:AbTasks.Contains($Id)) { throw "Unknown task '$Id'. Run 'abyte list' to see available tasks." }
    $script:AbTasks[$Id]
}

function Merge-AbParameter {
    param([hashtable]$Defaults, [hashtable]$Overrides)
    $p = @{}
    foreach ($k in $Defaults.Keys) { $p[$k] = $Defaults[$k] }
    if ($Overrides) {
        foreach ($k in $Overrides.Keys) {
            if (-not $Defaults.ContainsKey($k)) { throw "Unknown parameter '$k'. Allowed: $(($Defaults.Keys | Sort-Object) -join ', ')" }
            $p[$k] = $Overrides[$k]
        }
    }
    $p
}

function Get-AbExitCode {
    <#
    .SYNOPSIS
        Maps results to RMM exit codes: 1 if any fix failed, 3010 if a reboot is required, else 0.
        A check that could not run is reported but does not fail the job.
        With -FailOn Warning|Critical, findings at or above that severity also return 1 (for monitors).
    #>
    param([Parameter(Mandatory)][object[]]$Results, [ValidateSet('None', 'Warning', 'Critical')][string]$FailOn = 'None')
    if (@($Results | Where-Object { $_.Status -eq 'Failed' -and $_.Kind -eq 'Fix' }).Count) { return 1 }
    if ($FailOn -ne 'None') {
        $min = [array]::IndexOf($script:Severities, $FailOn)
        $hit = @($Results | ForEach-Object { $_.Findings } | Where-Object { [array]::IndexOf($script:Severities, $_.Severity) -ge $min })
        if ($hit.Count) { return 1 }
    }
    if (@($Results | Where-Object RebootRequired).Count) { return 3010 }
    0
}

function Get-AbVerdict {
    # Normalises a fix task's Test output to @{ Needed = [bool]; Reason = [string] }.
    # Tolerates stray pipeline output: the last hashtable wins.
    param([object[]]$Output)
    $h = @($Output) | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
    if (-not $h) { throw 'Fix task Test block must return @{ Needed = $true/$false; Reason = "..." }.' }
    @{ Needed = [bool]$h['Needed']; Reason = [string]$h['Reason'] }
}
