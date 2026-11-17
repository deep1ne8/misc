# Console styling: a pastel palette with three rendering modes.
#   TrueColor - 24-bit ANSI colours (Windows Terminal, modern consoles, PowerShell 7)
#   Basic     - the 16 console colours (older consoles)
#   None      - plain text (output redirected, NO_COLOR set, RMM logs)
# Glyphs are Unicode where the console can show them, ASCII otherwise. This file stays ASCII
# (Windows PowerShell 5.1 misreads BOM-less UTF-8), so glyphs are built from code points.

$script:AbPalette = [ordered]@{
    #           R    G    B   basic fallback
    Accent  = @(195, 177, 225, 'Magenta')     # lavender: titles, frames
    Ok      = @(168, 230, 207, 'Green')       # mint
    Warn    = @(255, 211, 182, 'Yellow')      # peach
    Bad     = @(255, 139, 148, 'Red')         # rose
    Info    = @(160, 210, 235, 'Cyan')        # sky
    Key     = @(253, 253, 150, 'Yellow')      # butter: menu keys, commands
    Soft    = @(224, 187, 228, 'DarkMagenta') # lilac: secondary headings
    Dim     = @(158, 158, 158, 'DarkGray')    # grey: hints and details
    Text    = @(235, 235, 235, 'Gray')
}

$script:AbStyle = $null

function Initialize-AbStyle {
    <#
      Chooses colour mode and glyphs. -Mode Auto detects; Color/Basic/None force a mode.
      Unicode glyphs need a UTF-8 capable console (Windows Terminal, VS Code, PowerShell 7).
    #>
    param([ValidateSet('Auto', 'TrueColor', 'Basic', 'None')][string]$Mode = 'Auto', [Nullable[bool]]$Unicode = $null)
    $redirected = $false
    try { $redirected = [Console]::IsOutputRedirected } catch { $redirected = $true }
    if ($Mode -eq 'Auto') {
        $vt = $false
        try { $vt = [bool]$Host.UI.SupportsVirtualTerminal } catch { $vt = $false }
        $Mode = if ($env:NO_COLOR -or $redirected) { 'None' }
        elseif ($vt -or $env:WT_SESSION -or $env:TERM_PROGRAM -or $env:COLORTERM) { 'TrueColor' }
        else { 'Basic' }
    }
    if ($null -eq $Unicode) {
        $cp = 0
        try { $cp = [Console]::OutputEncoding.CodePage } catch { $cp = 0 }
        $Unicode = (-not $redirected) -and ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode' -or $cp -eq 65001)
    }
    $u = [bool]$Unicode
    $g = if ($u) {
        @{ Ok = [string][char]0x2714; Bad = [string][char]0x2716; Warn = [string][char]0x25B2; Info = [string][char]0x2022; Partial = [string][char]0x25D0
            Skip = [string][char]0x2013; Arrow = [string][char]0x2192; Dot = [string][char]0x00B7; H = [string][char]0x2500; V = [string][char]0x2502
            TL = [string][char]0x256D; TR = [string][char]0x256E; BL = [string][char]0x2570; BR = [string][char]0x256F; Prompt = [string][char]0x203A }
    } else {
        @{ Ok = '+'; Bad = 'x'; Warn = '!'; Info = '-'; Partial = '~'; Skip = '-'; Arrow = '->'; Dot = '|'; H = '-'; V = '|'; TL = '+'; TR = '+'; BL = '+'; BR = '+'; Prompt = '>' }
    }
    $width = 100
    try { if (-not $redirected) { $width = [Math]::Max(60, [Math]::Min(110, $Host.UI.RawUI.WindowSize.Width - 2)) } } catch { $width = 100 }
    $script:AbStyle = [pscustomobject]@{ Mode = $Mode; Unicode = $u; Glyph = $g; Width = $width }
    $script:AbStyle
}

function Get-AbStyle { if (-not $script:AbStyle) { Initialize-AbStyle | Out-Null }; $script:AbStyle }

function Write-AbLine {
    <#
      Writes one line made of (text, role) pairs, e.g.
        Write-AbLine '  ' Text 'OK' Ok ' all good' Dim
      Roles: Accent Ok Warn Bad Info Key Soft Dim Text, optionally with '+' for bold (e.g. 'Accent+').
      Display-only (Write-Host); data goes to the pipeline elsewhere.
    #>
    param([Parameter(ValueFromRemainingArguments)][object[]]$Parts)
    $st = Get-AbStyle
    $esc = [char]27
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Parts.Count; $i += 2) {
        $text = [string]$Parts[$i]
        $role = if ($i + 1 -lt $Parts.Count) { [string]$Parts[$i + 1] } else { 'Text' }
        $bold = $role.EndsWith('+'); $role = $role.TrimEnd('+')
        if (-not $script:AbPalette.Contains($role)) { $role = 'Text' }
        $c = $script:AbPalette[$role]
        switch ($st.Mode) {
            'TrueColor' { [void]$sb.Append(('{0}[{1}38;2;{2};{3};{4}m{5}{0}[0m' -f $esc, $(if ($bold) { '1;' } else { '' }), $c[0], $c[1], $c[2], $text)) }
            'Basic' { Write-Host $text -NoNewline -ForegroundColor $c[3] }
            default { [void]$sb.Append($text) }
        }
    }
    if ($st.Mode -eq 'Basic') { Write-Host '' } else { Write-Host $sb.ToString() }
}

function Format-AbWrap {
    # Word-wraps text to the console width; returns lines (without the indent).
    param([string]$Text, [int]$Indent = 0)
    $max = [Math]::Max(30, (Get-AbStyle).Width - $Indent)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($para in ([string]$Text -split "`r?`n")) {
        $line = ''
        foreach ($word in ($para -split ' ')) {
            if ($line.Length -and ($line.Length + 1 + $word.Length) -gt $max) { $lines.Add($line); $line = $word }
            elseif ($line.Length) { $line += " $word" } else { $line = $word }
        }
        $lines.Add($line)
    }
    $lines
}

function Format-AbDuration {
    param([double]$Milliseconds)
    $t = [TimeSpan]::FromMilliseconds($Milliseconds)
    if ($t.TotalSeconds -lt 1) { return '<1 s' }
    if ($t.TotalMinutes -lt 1) { return ('{0:0.#} s' -f $t.TotalSeconds) }
    if ($t.TotalHours -lt 1) { return ('{0} min {1} s' -f [int][Math]::Floor($t.TotalMinutes), $t.Seconds) }
    ('{0} h {1} min' -f [int][Math]::Floor($t.TotalHours), $t.Minutes)
}
