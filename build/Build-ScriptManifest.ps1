<#
.SYNOPSIS
    Generates scripts.manifest.json: the single source of truth for the AutoByte script library.

.DESCRIPTION
    Walks Scripts/ and records, for every non-empty .ps1/.sh/.cmd file:
      name, path, sha256, version, description, mutating, requiresAdmin,
      interactive, unsupported, agentDiagnostic.

    sha256 is computed over the file content with CRLF normalised to LF, so the hash
    is identical for a Windows checkout, a Linux checkout, and raw.githubusercontent.com.
    Consumers (AutoByte.ps1, AutoByteGUI.py, AutoByte-Agent) must normalise the same way.

    'mutating' is detected from the PowerShell AST and is conservative: any command
    that can change state marks the script mutating. Reviewed exceptions live in
    build/manifest.overrides.json with a reason.

.PARAMETER Root
    Repository root. Defaults to the parent of this script's folder.

.PARAMETER Check
    Do not write; exit 1 if the committed manifest is stale. Used by CI.

.EXAMPLE
    pwsh ./build/Build-ScriptManifest.ps1
.EXAMPLE
    pwsh ./build/Build-ScriptManifest.ps1 -Check

.NOTES
    Version: 1.0.0
    Exit codes: 0 ok, 1 stale manifest / error.
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [switch]$Check
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ManifestPath = Join-Path $Root 'scripts.manifest.json'
$OverridesPath = Join-Path $PSScriptRoot 'manifest.overrides.json'

# Cmdlets that only read, format, log, or build objects. Everything else with a
# state-changing verb is treated as mutating.
$SafeCommands = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@(
    'Start-Transcript', 'Stop-Transcript', 'Start-Sleep', 'Add-Type', 'Add-Member', 'New-Object',
    'New-TimeSpan', 'New-CimSession', 'Remove-CimSession', 'New-PSSessionOption', 'New-CimSessionOption',
    'Import-Module', 'Remove-Module', 'Export-Csv', 'Export-Clixml', 'Out-File', 'Add-Content',
    'Set-StrictMode', 'Set-Variable', 'New-Variable', 'Remove-Variable', 'Clear-Variable', 'Clear-Host',
    'Set-Location', 'Push-Location', 'Pop-Location', 'Start-Job', 'Stop-Job', 'Remove-Job',
    'Update-TypeData', 'Update-FormatData', 'New-Guid', 'New-Event', 'Register-EngineEvent',
    'Set-PSDebug', 'Enable-PSBreakpoint', 'Disable-PSBreakpoint', 'New-Alias', 'Set-Alias'
) | ForEach-Object { [void]$SafeCommands.Add($_) }

$MutatingVerbs = 'Set|New|Remove|Install|Uninstall|Update|Clear|Add|Enable|Disable|Reset|Register|Unregister|Start|Stop|Restart|Suspend|Resume|Rename|Move|Copy|Repair|Grant|Revoke|Deny|Block|Unblock|Initialize|Format|Mount|Dismount|Publish|Unpublish|Expand|Compress|Send|Save|Sync|Invoke|Write-EventLog|Connect|Disconnect|Lock|Unlock|Protect|Unprotect|Optimize|Edit|Merge|Join|Split|Submit|Approve|Deploy|Request|Import'
$MutatingRegex = [regex]"^(?:$MutatingVerbs)-"
$NativeMutators = 'sfc|dism|msiexec|reg|net|net1|sc|winget|choco|bcdedit|cleanmgr|cipher|icacls|takeown|wusa|usoclient|schtasks|netsh|ipconfig|shutdown|powercfg|vssadmin|diskpart|format|del|erase|rd|rmdir|robocopy|xcopy|copy|move|wmic|setup|wuauclt|gpupdate|certutil|regsvr32|rundll32|bitsadmin|pnputil|manage-bde|chkdsk|defrag|attrib|cmd|powershell|pwsh|curl|wget|tar|expand|mkdir|md|cscript|wscript|taskkill|logoff|klist|sudo|rm|softwareupdate|installer|launchctl|defaults|diskutil|tmutil|chmod|chown|mv|cp|ln|pkgutil|profiles|systemsetup|spctl|xattr|killall|pkill|open|dscl|csrutil'
$NativeRegex = [regex]"^(?:$NativeMutators)(?:\.exe)?$"

function Get-NormalizedHash {
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    # Hash raw bytes (BOM included) with CRLF->LF only.
    $ms = [System.IO.MemoryStream]::new()
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 13 -and ($i + 1) -lt $bytes.Length -and $bytes[$i + 1] -eq 10) { continue }
        $ms.WriteByte($bytes[$i])
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($ms.ToArray())) -replace '-', '').ToLowerInvariant() }
    finally { $sha.Dispose(); $ms.Dispose() }
}

function Get-ScriptDescription {
    param([string]$Text)
    if ($Text -match '(?ms)^\s*\.SYNOPSIS\s*\r?\n(.*?)(?:\r?\n\s*\r?\n|\r?\n\s*\.[A-Z]+|#>)') {
        $line = ($Matches[1] -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 1)
        if ($line) { return $line }
    }
    foreach ($l in $Text -split '\r?\n') {
        $t = $l.Trim()
        if ($t -match '^(#!|<#|#>|#\s*-+|#\s*=+|#Requires)' -or $t -eq '#' -or $t -eq '') { continue }
        if ($t -match '^(#|::|REM\s)\s*(.+)$') { return $Matches[2].Trim() }
        break
    }
    return ''
}

function Get-PowerShellTraits {
    param([string]$Path)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    $reasons = [System.Collections.Generic.List[string]]::new()
    $interactive = $false
    $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($c in $commands) {
        $name = $c.GetCommandName()
        if (-not $name) { $reasons.Add('dynamic-invocation'); continue }
        $leaf = [System.IO.Path]::GetFileName($name)
        if ($leaf -match '^(Read-Host|Get-Credential|Out-GridView)$') { $interactive = $true; continue }
        if ($SafeCommands.Contains($leaf)) { continue }
        if ($leaf -match '^(iex|Invoke-Expression|Invoke-Command|Invoke-Item|Invoke-WmiMethod|Invoke-CimMethod|Invoke-RestMethod|Invoke-WebRequest)$') {
            # Web calls are read-only unless they write a file or POST/PUT/DELETE.
            if ($leaf -match 'Invoke-(RestMethod|WebRequest)') {
                $txt = $c.Extent.Text
                if ($txt -notmatch '(?i)-OutFile|-Method\s+[''"]?(Post|Put|Delete|Patch)') { continue }
            }
            $reasons.Add($leaf); continue
        }
        if ($MutatingRegex.IsMatch($leaf) -or $NativeRegex.IsMatch($leaf)) { $reasons.Add($leaf) }
    }
    # .NET/COM calls that change state
    $members = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)
    foreach ($m in $members) {
        $mn = "$($m.Member)"
        if ($mn -match '^(Delete|DeleteSubKey\w*|SetValue|CreateSubKey|Kill|Start|Install|Uninstall|Put|Create|DownloadFile|WriteAllText|WriteAllBytes|Copy|Move|SetAccessControl|SetInfo|Commit|Save)$') { $reasons.Add(".$mn()") }
    }
    # Property writes on objects (e.g. $outlook.Session.DefaultStore = ...) can change COM/WMI state.
    $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
    foreach ($a in $assigns) {
        $target = $a.Left.Expression
        while ($target -is [System.Management.Automation.Language.MemberExpressionAst]) { $target = $target.Expression }
        $isLocalObject = $target -is [System.Management.Automation.Language.VariableExpressionAst] -and $a.Left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $ast.Extent.Text -match ('\$' + [regex]::Escape($target.VariablePath.UserPath) + '\s*=\s*(\[pscustomobject\]|\[ordered\]|@\{|New-Object\s+(PSObject|psobject|System\.Collections))')
        if (-not $isLocalObject) { $reasons.Add('property-assignment') }
    }
    foreach ($c in $commands) {
        if ($c.GetCommandName() -eq 'New-Object' -and $c.Extent.Text -match '(?i)-ComObject') { $reasons.Add('COM-automation') }
    }
    if (@($members | Where-Object { "$($_.Member)" -match '^(Popup|ShowDialog|ReadKey|ReadLine)$' }).Count -gt 0) { $interactive = $true }
    $text = Get-Content -LiteralPath $Path -Raw
    $explicitAdmin = ($text -match '(?im)^#Requires\s+-RunAsAdministrator') -or
    ($text -match 'WindowsBuiltInRole\]\s*(::)?\s*[''"]?Administrator') -or
    ($text -match 'IsInRole\(\s*[''"]?Administrators?')
    $systemWrites = ($reasons.Count -gt 0) -and (
        ($text -match '(?i)(Set|New|Remove)-ItemProperty[^\r\n]*HKLM') -or
        (@($reasons | Where-Object { $_ -match '^(sfc|dism|msiexec|bcdedit|wusa|usoclient|vssadmin|diskpart|pnputil|Install-|Uninstall-|Restart-Service|Stop-Service|Set-Service)' }).Count -gt 0))
    $requiresAdmin = $explicitAdmin -or $systemWrites
    [pscustomobject]@{
        ParseErrors   = @($errors).Count
        Mutating      = $reasons.Count -gt 0
        Reasons       = @($reasons | Select-Object -Unique)
        Interactive   = $interactive
        RequiresAdmin = [bool]$requiresAdmin
        Text          = $text
    }
}

$overrides = @{}
if (Test-Path -LiteralPath $OverridesPath) {
    $raw = Get-Content -LiteralPath $OverridesPath -Raw | ConvertFrom-Json
    foreach ($p in $raw.scripts.PSObject.Properties) { $overrides[$p.Name] = $p.Value }
}

$scriptsDir = Join-Path $Root 'Scripts'
$files = Get-ChildItem -LiteralPath $scriptsDir -Recurse -File |
    Where-Object { $_.Extension -in '.ps1', '.sh', '.cmd' -and $_.Length -gt 0 } |
    Sort-Object { $_.FullName.Substring($Root.Length).Replace('\', '/') } -Culture ([cultureinfo]::InvariantCulture)

$entries = foreach ($f in $files) {
    $rel = $f.FullName.Substring($Root.Length).TrimStart('\', '/').Replace('\', '/')
    if ($f.Extension -eq '.ps1') {
        $t = Get-PowerShellTraits -Path $f.FullName
        if ($t.ParseErrors -gt 0) { throw "Parse errors in $rel; fix before generating the manifest." }
    }
    else {
        $text = Get-Content -LiteralPath $f.FullName -Raw
        $t = [pscustomobject]@{
            Mutating = $true; Reasons = @('non-powershell'); Interactive = ($text -match '\bread\s+-|set\s+/p')
            RequiresAdmin = ($text -match 'sudo|EUID|net session|Administrator'); Text = $text
        }
    }
    $version = if ($t.Text -match '(?im)^\s*(?:#\s*)?Version\s*[:=]\s*v?([0-9]+(?:\.[0-9]+){1,3})') { $Matches[1] } else { '0.0.0' }
    $platform = if ($f.Extension -eq '.sh') { 'macos' } else { 'windows' }
    $unsupported = $f.BaseName -match '(?i)bypass|skip_?tpm'

    $entry = [ordered]@{
        name            = $f.BaseName
        path            = $rel
        sha256          = Get-NormalizedHash -Path $f.FullName
        version         = $version
        description     = Get-ScriptDescription -Text $t.Text
        platform        = $platform
        mutating        = [bool]$t.Mutating
        requiresAdmin   = [bool]$t.RequiresAdmin
        interactive     = [bool]$t.Interactive
        unsupported     = [bool]$unsupported
        agentDiagnostic = $false
    }
    if ($overrides.ContainsKey($f.BaseName)) {
        $o = $overrides[$f.BaseName]
        foreach ($k in 'mutating', 'requiresAdmin', 'interactive', 'unsupported', 'description') {
            if ($o.PSObject.Properties[$k]) { $entry[$k] = $o.$k }
        }
    }
    # Only non-mutating, non-interactive, supported Windows scripts may auto-run in the agent.
    $entry.agentDiagnostic = (-not $entry.mutating) -and (-not $entry.interactive) -and (-not $entry.unsupported) -and $platform -eq 'windows'
    [pscustomobject]$entry
}

$manifest = [ordered]@{
    schemaVersion = 1
    hashAlgorithm = 'sha256-lf'
    scriptCount   = @($entries).Count
    scripts       = @($entries)
}
$json = ($manifest | ConvertTo-Json -Depth 5) -replace "`r`n", "`n"
$json += "`n"

if ($Check) {
    $current = if (Test-Path -LiteralPath $ManifestPath) { (Get-Content -LiteralPath $ManifestPath -Raw) -replace "`r`n", "`n" } else { '' }
    if ($current -ne $json) {
        Write-Error 'scripts.manifest.json is stale. Run: pwsh ./build/Build-ScriptManifest.ps1 and commit the result.' -ErrorAction Continue
        exit 1
    }
    Write-Output "Manifest up to date ($($manifest.scriptCount) scripts)."
    exit 0
}

if ($PSCmdlet.ShouldProcess($ManifestPath, 'Write manifest')) {
    [System.IO.File]::WriteAllText($ManifestPath, $json, [System.Text.UTF8Encoding]::new($false))
    Write-Output "Wrote $ManifestPath ($($manifest.scriptCount) scripts; $(@($entries | Where-Object agentDiagnostic).Count) agent diagnostics)."
}
exit 0
