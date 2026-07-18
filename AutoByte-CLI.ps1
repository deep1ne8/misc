<#
 .SYNOPSIS
    AutoByte CLI - a self-discovering launcher for the AutoByte script library.
    Lists every script in ./Scripts, shows what it does, and runs the one you pick.

 .USAGE
    AutoByte-CLI.ps1              # interactive numbered menu
    AutoByte-CLI.ps1 list         # print all scripts with descriptions
    AutoByte-CLI.ps1 run <name>   # run a script by filename (tab-free)
    AutoByte-CLI.ps1 show <name>  # print a script's header/description
    AutoByte-CLI.ps1 categories   # list scripts grouped by inferred category
#>
# AutoByte CLI - self-discovering launcher. Reads args from $args directly.
$Command = if ($args.Count -ge 1) { $args[0] } else { '' }
$Name    = if ($args.Count -ge 2) { $args[1] } else { '' }

$ErrorActionPreference = 'Stop'
$scriptsDir = Join-Path $PSScriptRoot 'Scripts'
if (-not (Test-Path $scriptsDir)) { Write-Host "Scripts dir not found at $scriptsDir" -ForegroundColor Red; exit 1 }

# ---- discover + describe ----------------------------------------------------
function Get-Description { param([string]$Path)
    $txt = Get-Content $Path -Raw -ErrorAction SilentlyContinue
    if ($null -eq $txt) { return '(no description)' }
    # prefer a .SYNOPSIS / .DESCRIPTION block
    if ($txt -match '(?s)\.SYNOPSIS\s*(.*?)(\.|\Z)') { return (($Matches[1].Trim() -split "`n") | Where-Object { $_.Trim() } | Select-Object -First 1).Trim() }
    # else first <# ... #> comment
    if ($txt -match '(?s)<#\s*(.*?)#>') {
        $c = (($Matches[1].Trim() -split "`n") | Where-Object { $_.Trim() } | Select-Object -First 1)
        if ($c) { return $c.Trim() }
    }
    # else first # comment line
    $line = (($txt -split "`n") | Where-Object { $_ -match '^\s*#' } | Select-Object -First 1)
    if ($line) { return ($line -replace '^\s*#\s?', '').Trim() }
    return '(no description)'
}

function Get-Category { param([string]$File)
    $f = $File.ToLower()
    switch -Regex ($f) {
        'printer|toner|suppl'                { return 'Printers & Supplies' }
        'office|teams|m365|excel|outlook'    { return 'Office / M365' }
        'repair|update|upgrade|windowsonline|windowssystem|windowos|cbs|searchdb|restart|feature' { return 'Windows Repair & Updates' }
        'network|internet|latency|speed|wlog|getwindowsevents|splitlog|httpserver|netscan' { return 'Network & Diagnostics' }
        'user|profile|admin|onedrive|entra|home' { return 'User / Profile / Admin' }
        'dell|hp|raid|certkey|bloat'         { return 'Dell / HP Hardware' }
        'calendar|exchange|smtp|3cx|gsuite|shared' { return 'Calendar / Exchange' }
        'install|download|winget|fido|revit|msi|deploy|certchain|financial|sage|qb|workpaper|cch|build' { return 'Deployment / Install' }
        'language|diskclean|checkdrive|memory|intune|debug|test|scriptstar' { return 'Utilities' }
        default                             { return 'General' }
    }
}

$all = Get-ChildItem $scriptsDir -Filter *.ps1 | ForEach-Object {
    [pscustomobject]@{
        file     = $_.Name
        name     = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        category = Get-Category $_.Name
        desc     = Get-Description $_.FullName
        path     = $_.FullName
    }
} | Sort-Object category, name

function Show-Table { param($Items)
    $w = ($Items.file | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    if (-not $w) { $w = 30 }
    foreach ($it in $Items) {
        Write-Host ("  {0,-$w}  {1}" -f $it.file, $it.desc) -ForegroundColor DarkGray
    }
}

# ---- commands ---------------------------------------------------------------
switch ($Command) {
    'list' {
        Write-Host "`nAutoByte scripts ($($all.Count)):`n" -ForegroundColor Cyan
        Show-Table $all
        Write-Host ''
    }
    'categories' {
        Write-Host "`nAutoByte scripts by category:`n" -ForegroundColor Cyan
        $all | Group-Object category | Sort-Object Name | ForEach-Object {
            Write-Host "  $($_.Name) ($($_.Count))" -ForegroundColor Yellow
            Show-Table $_.Group
            Write-Host ''
        }
    }
    'show' {
        if (-not $Name) { Write-Host 'usage: AutoByte-CLI.ps1 show <name>' -ForegroundColor Yellow; exit 1 }
        $m = $all | Where-Object { $_.name -eq $Name -or $_.file -eq $Name }
        if (-not $m) { Write-Host "not found: $Name" -ForegroundColor Red; exit 1 }
        Write-Host "`n== $($m.file) ==" -ForegroundColor Cyan
        Write-Host $m.desc -ForegroundColor White
        Write-Host "Category: $($m.category)`n" -ForegroundColor DarkGray
    }
    'run' {
        if (-not $Name) { Write-Host 'usage: AutoByte-CLI.ps1 run <name>' -ForegroundColor Yellow; exit 1 }
        $m = $all | Where-Object { $_.name -eq $Name -or $_.file -eq $Name }
        if (-not $m) { Write-Host "not found: $Name" -ForegroundColor Red; exit 1 }
        Write-Host "`nRunning $($m.file) ...`n" -ForegroundColor Green
        & powershell -NoProfile -ExecutionPolicy Bypass -File $m.path
    }
    '' {
        # interactive menu
        Write-Host "`nAutoByte CLI - pick a script to run`n" -ForegroundColor Cyan
        $i = 1; $menu = @()
        foreach ($it in $all) {
            Write-Host ("  [{0,3}] {1,-42} {2}" -f $i, $it.file, $it.desc) -ForegroundColor DarkGray
            $menu += $it; $i++
        }
        Write-Host ''
        $pick = Read-Host 'Enter number (or q to quit)'
        if ($pick -match '^q$') { return }
        if ($pick -notmatch '^\d+$' -or [int]$pick -lt 1 -or [int]$pick -gt $menu.Count) {
            Write-Host 'Invalid selection.' -ForegroundColor Red; exit 1
        }
        $sel = $menu[[int]$pick - 1]
        Write-Host "`nRunning $($sel.file) ...`n" -ForegroundColor Green
        & powershell -NoProfile -ExecutionPolicy Bypass -File $sel.path
    }
    default {
        Write-Host "Unknown command: $Command" -ForegroundColor Red
        Write-Host 'Usage: AutoByte-CLI.ps1 [list|categories|show <name>|run <name>]' -ForegroundColor Yellow
        exit 1
    }
}
