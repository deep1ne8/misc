<#
.SYNOPSIS
    Fails if any file contains a credential that previously leaked from this repository.

.DESCRIPTION
    build/leaked-secrets.sha256 lists the SHA-256, length and label of each leaked credential.
    The values themselves are never stored. This script finds every run of token characters
    (letters, digits, + / = % _ -) in the scanned text, hashes every substring whose length matches
    a listed entry, and reports file, line and label (never the value) for each match.

    Use it to stop a rotated-but-still-valid key, or a copy-pasted old script, from being committed again.

.PARAMETER Root
    Repository root. Defaults to the parent of this script's folder.

.PARAMETER Staged
    Scan the staged (index) version of changed files instead of the working tree. Used by the pre-commit hook.

.PARAMETER Path
    Scan only these files (relative to Root). Default: every file tracked by git.

.EXAMPLE
    pwsh ./build/Test-LeakedSecrets.ps1
.EXAMPLE
    pwsh ./build/Test-LeakedSecrets.ps1 -Staged

.NOTES
    Version: 1.0.0
    Exit codes: 0 clean, 1 leaked credential found or error.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [switch]$Staged,
    [string[]]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$listFile = Join-Path $PSScriptRoot 'leaked-secrets.sha256'
$entries = @{}
foreach ($line in Get-Content -LiteralPath $listFile) {
    if ($line -match '^\s*#' -or -not $line.Trim()) { continue }
    $parts = $line.Trim() -split '\s+'
    if ($parts.Count -lt 3 -or $parts[0] -notmatch '^[0-9a-f]{64}$') { throw "Bad line in ${listFile}: $line" }
    $entries[$parts[0]] = [pscustomobject]@{ Length = [int]$parts[1]; Label = $parts[2] }
}
$lengths = @($entries.Values.Length | Sort-Object -Unique)
$minLen = ($lengths | Measure-Object -Minimum).Minimum

Push-Location -LiteralPath $Root
try {
    if (-not $Path) {
        $Path = if ($Staged) { @(& git diff --cached --name-only --diff-filter=ACMR) } else { @(& git ls-files) }
        if ($LASTEXITCODE -ne 0) { throw 'git failed to list files.' }
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $utf8 = [System.Text.Encoding]::UTF8
    $tokenRx = [regex]"[A-Za-z0-9+/=%_\-]{$minLen,}"
    $findings = [System.Collections.Generic.List[object]]::new()
    $self = 'build/leaked-secrets.sha256'

    foreach ($file in $Path) {
        if ($file -eq $self) { continue }
        $text = if ($Staged) { (& git show ":$file") -join "`n" } else {
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }
            $bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $file).Path)
            if ($bytes -contains 0) { continue }  # binary
            $utf8.GetString($bytes)
        }
        if (-not $text) { continue }
        $lineNo = 1; $pos = 0
        foreach ($m in $tokenRx.Matches($text)) {
            $tok = $m.Value
            foreach ($len in $lengths) {
                if ($len -gt $tok.Length) { break }
                for ($i = 0; $i -le $tok.Length - $len; $i++) {
                    $h = ([BitConverter]::ToString($sha.ComputeHash($utf8.GetBytes($tok.Substring($i, $len)))) -replace '-', '').ToLowerInvariant()
                    if ($entries.ContainsKey($h)) {
                        $lineNo += ([regex]::Matches($text.Substring($pos, $m.Index - $pos), "`n")).Count; $pos = $m.Index
                        $findings.Add([pscustomobject]@{ File = $file; Line = $lineNo; Label = $entries[$h].Label })
                    }
                }
            }
        }
    }
    $sha.Dispose()
}
finally { Pop-Location }

if ($findings.Count) {
    foreach ($f in $findings) { Write-Output "LEAKED CREDENTIAL: $($f.File):$($f.Line)  [$($f.Label)]" }
    Write-Error "$($findings.Count) previously leaked credential(s) found. Remove them; use parameters or environment variables instead." -ErrorAction Continue
    exit 1
}
Write-Output "No leaked credentials found ($(@($Path).Count) files, $($entries.Count) fingerprints)."
exit 0
