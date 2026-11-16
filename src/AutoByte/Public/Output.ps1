function ConvertTo-AbJson {
    <#
    .SYNOPSIS
        Serialises scan/task results to JSON for RMM and automation (stable field names).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)]$InputObject, [switch]$Compress)
    begin { $items = New-Object System.Collections.Generic.List[object] }
    process { foreach ($o in @($InputObject)) { $items.Add($o) } }
    end {
        # One object -> JSON object; several -> one JSON array (never several concatenated documents).
        if ($items.Count -eq 1) { ConvertTo-Json -InputObject $items[0] -Depth 8 -Compress:$Compress }
        else { ConvertTo-Json -InputObject @($items.ToArray()) -Depth 8 -Compress:$Compress }
    }
}

function ConvertTo-AbHtmlText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Export-AbReport {
    <#
    .SYNOPSIS
        Writes a self-contained HTML health report (no external files or scripts) for a scan.
    .PARAMETER Scan
        Output of Invoke-AbScan.
    .PARAMETER FixResults
        Optional output of Invoke-AbFix, shown in an "Actions taken" section.
    .PARAMETER Path
        Output file. Default: %ProgramData%\AutoByte\Reports\<Computer>_<timestamp>.html
    .EXAMPLE
        Invoke-AbScan | Export-AbReport -Open
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]$Scan,
        [object[]]$FixResults = @(),
        [string]$Path,
        [string]$Organization = $env:AB_ORGANIZATION,
        [switch]$Open
    )
    process {
        if (-not $Path) {
            $base = if ($env:ProgramData) { $env:ProgramData } else { [IO.Path]::GetTempPath() }
            $Path = Join-Path $base ('AutoByte\Reports\{0}_{1}.html' -f $Scan.Computer, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        }
        $e = { param($v) ConvertTo-AbHtmlText $v }
        $sev = @{ Critical = 3; Warning = 2; Info = 1 }
        $overall = if ($Scan.Critical) { 'Needs attention' } elseif ($Scan.Warning) { 'Mostly healthy' } else { 'Healthy' }
        $overallClass = if ($Scan.Critical) { 'crit' } elseif ($Scan.Warning) { 'warn' } else { 'ok' }

        $rows = foreach ($f in ($Scan.Findings | Sort-Object { - $sev[$_.Severity] }, Title)) {
            $cls = $f.Severity.ToLowerInvariant().Substring(0, 4)
            $fix = if ($f.FixTask) { "<code>abyte fix $(& $e $f.FixTask)</code>" } else { '' }
            "<tr><td><span class=""pill $cls"">$(& $e $f.Severity)</span></td><td><b>$(& $e $f.Title)</b><div class=""muted"">$(& $e $f.Detail)</div></td><td>$fix</td></tr>"
        }
        if (-not $rows) { $rows = '<tr><td colspan="3">No problems found.</td></tr>' }

        $checks = foreach ($r in $Scan.Results) {
            $name = (Get-AbTask -Id $r.TaskId).Name
            "<tr><td>$(& $e $name)</td><td>$(& $e $r.Status)</td><td class=""muted"">$($r.DurationMs) ms</td></tr>"
        }
        $actions = foreach ($r in $FixResults) {
            $name = (Get-AbTask -Id $r.TaskId).Name
            "<tr><td>$(& $e $name)</td><td>$(& $e $r.Status)</td><td class=""muted"">$(& $e (($r.Messages + $r.Error) -join ' '))</td></tr>"
        }
        $sum = $Scan.Summary
        $html = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>AutoByte report - $(& $e $Scan.Computer)</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f6f7f9;color:#1f2328}
main{max-width:960px;margin:0 auto;padding:24px}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:28px 0 8px}
.card{background:#fff;border:1px solid #d8dee4;border-radius:10px;padding:16px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px}
.k{font-size:12px;color:#59636e}.v{font-size:15px;font-weight:600}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d8dee4;border-radius:10px;overflow:hidden}
td{padding:10px 12px;border-top:1px solid #eef1f4;vertical-align:top;font-size:14px}
.muted{color:#59636e;font-size:13px}.pill{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600}
.crit{background:#ffebe9;color:#a40e26}.warn{background:#fff8c5;color:#7d4e00}.info{background:#ddf4ff;color:#0a3069}.ok{background:#dafbe1;color:#116329}
code{background:#eef1f4;padding:2px 6px;border-radius:4px;font-size:12px}
footer{margin-top:28px;font-size:12px;color:#59636e}
@media print{body{background:#fff}main{padding:0}}
</style></head><body><main>
<h1>PC health report: $(& $e $Scan.Computer)</h1>
<div class="muted">$(& $e $Organization) $(if ($Organization) { '&middot;' }) Generated $(& $e (Get-Date -Format 'yyyy-MM-dd HH:mm'))</div>
<h2>Overall</h2>
<div class="card grid">
<div><div class="k">Status</div><div class="v"><span class="pill $overallClass">$overall</span></div></div>
<div><div class="k">Critical</div><div class="v">$($Scan.Critical)</div></div>
<div><div class="k">Warnings</div><div class="v">$($Scan.Warning)</div></div>
<div><div class="k">Checks run</div><div class="v">$(@($Scan.Results).Count)</div></div>
</div>
<h2>Computer</h2>
<div class="card grid">
<div><div class="k">Windows</div><div class="v">$(& $e $sum.OS)</div><div class="muted">$(& $e $sum.Version)</div></div>
<div><div class="k">Model</div><div class="v">$(& $e $sum.Manufacturer) $(& $e $sum.Model)</div></div>
<div><div class="k">Memory</div><div class="v">$(& $e $sum.RamGB) GB</div></div>
<div><div class="k">Last restart</div><div class="v">$(& $e $sum.LastBoot)</div></div>
</div>
<h2>Findings</h2>
<table>$($rows -join "`n")</table>
$(if ($actions) { "<h2>Actions taken</h2><table>$($actions -join "`n")</table>" })
<h2>Checks</h2>
<table>$($checks -join "`n")</table>
<footer>AutoByte $(& $e $sum.AutoByte). Checks are read-only; fixes run only when requested.</footer>
</main></body></html>
"@
        if ($PSCmdlet.ShouldProcess($Path, 'Write HTML report')) {
            $dir = Split-Path -Parent $Path
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($true)))
            Write-AbLog -Message "Report written: $Path"
            if ($Open) { Start-Process -FilePath $Path }
        }
        Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    }
}

function Invoke-AbLegacyScript {
    <#
    .SYNOPSIS
        Runs a v2 library script (Scripts/*.ps1) through the hash-verifying v2 launcher, so existing
        tools keep working while they are migrated to tasks.
    .EXAMPLE
        Invoke-AbLegacyScript -Name DiskCleaner -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_.\-]+$')][string]$Name,
        [ValidatePattern('^(v\d+\.\d+\.\d+|[0-9a-f]{40})$')][string]$Ref = 'v2.0.0'
    )
    $launcherUrl = "https://raw.githubusercontent.com/deep1ne8/misc/$Ref/AutoByte.ps1"
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("AutoByte-launcher-{0}.ps1" -f [guid]::NewGuid().ToString('N'))
    try {
        Invoke-WebRequest -Uri $launcherUrl -OutFile $tmp -UseBasicParsing
        $exe = (Get-Process -Id $PID).Path
        $launcherArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $tmp, '-Ref', $Ref, '-Script', $Name)
        if ($WhatIfPreference) { $launcherArgs += '-WhatIf' }
        & $exe @launcherArgs
        $LASTEXITCODE
    }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}
