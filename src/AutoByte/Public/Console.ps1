function Initialize-AbConsole {
    <#
    .SYNOPSIS
        Sets how AutoByte draws on the console: Auto (default), TrueColor (pastel 24-bit colours),
        Basic (16 colours) or None (plain text). -Ascii forces ASCII glyphs.
    #>
    [CmdletBinding()]
    param([ValidateSet('Auto', 'TrueColor', 'Basic', 'None')][string]$Mode = 'Auto', [switch]$Ascii)
    $u = if ($Ascii) { $false } else { $null }
    Initialize-AbStyle -Mode $Mode -Unicode $u
}

function Get-AbStatusView {
    # Plain-language label, colour role and glyph for a result status.
    param([Parameter(Mandatory)][string]$Status, [string]$Kind = 'Fix')
    $g = (Get-AbStyle).Glyph
    switch ($Status) {
        'Healthy' { @{ Label = 'Healthy'; Role = 'Ok'; Glyph = $g.Ok } }
        'Issues' { @{ Label = 'Needs attention'; Role = 'Warn'; Glyph = $g.Warn } }
        'NotNeeded' { @{ Label = 'Nothing to do'; Role = 'Ok'; Glyph = $g.Ok } }
        'Fixed' { @{ Label = 'Fixed'; Role = 'Ok'; Glyph = $g.Ok } }
        'Partial' { @{ Label = 'Partly fixed'; Role = 'Warn'; Glyph = $g.Partial } }
        'Failed' { @{ Label = $(if ($Kind -eq 'Check') { 'Could not run' } else { 'Not fixed' }); Role = 'Bad'; Glyph = $g.Bad } }
        'Skipped' { @{ Label = 'Skipped'; Role = 'Dim'; Glyph = $g.Skip } }
        'WhatIf' { @{ Label = 'Preview only'; Role = 'Info'; Glyph = $g.Info } }
        default { @{ Label = $Status; Role = 'Text'; Glyph = $g.Info } }
    }
}

function Write-AbRule {
    param([string]$Title = '')
    $st = Get-AbStyle
    $w = [Math]::Min(76, $st.Width - 2)
    if ($Title) {
        $line = $st.Glyph.H * [Math]::Max(3, $w - $Title.Length - 4)
        Write-AbLine '  ' Text $Title 'Accent+' ' ' Text $line Dim
    }
    else { Write-AbLine '  ' Text ($st.Glyph.H * $w) Dim }
}

function Write-AbBox {
    <#
    .SYNOPSIS
        Draws a rounded box. Each line is an array of (text, role) pairs.
    #>
    param([Parameter(Mandatory)][object[]]$Lines, [int]$Width = 0)
    $st = Get-AbStyle; $g = $st.Glyph
    $Lines = @(foreach ($l in $Lines) { if (@($l).Count -eq 1 -and @($l)[0] -is [array]) { , @(@($l)[0]) } else { , @($l) } })
    $visible = @(foreach ($l in $Lines) { $n = 0; for ($i = 0; $i -lt $l.Count; $i += 2) { $n += ([string]$l[$i]).Length }; $n })
    $inner = [Math]::Max(($visible | Measure-Object -Maximum).Maximum + 4, $Width)
    Write-AbLine '  ' Text ($g.TL + ($g.H * $inner) + $g.TR) Accent
    for ($k = 0; $k -lt $Lines.Count; $k++) {
        $parts = @('  ', 'Text', $g.V, 'Accent', '  ', 'Text') + @($Lines[$k]) + @((' ' * ($inner - 2 - $visible[$k])), 'Text', $g.V, 'Accent')
        Write-AbLine @parts
    }
    Write-AbLine '  ' Text ($g.BL + ($g.H * $inner) + $g.BR) Accent
}

function Write-AbWrapped {
    # Writes wrapped text at an indent in one colour role.
    param([string]$Text, [int]$Indent = 5, [string]$Role = 'Dim', [string]$FirstPrefix = '')
    $pad = ' ' * $Indent
    $first = $true
    foreach ($l in (Format-AbWrap -Text $Text -Indent ($Indent + $FirstPrefix.Length))) {
        if ($first -and $FirstPrefix) { Write-AbLine $pad Text $FirstPrefix $Role $l $Role } else { Write-AbLine ($pad + (' ' * $FirstPrefix.Length)) Text $l $Role }
        $first = $false
    }
}

function Write-AbScanView {
    <#
    .SYNOPSIS
        Shows a scan (Invoke-AbScan output) as a tidy, colour-coded summary for people.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)]$Scan, [switch]$EndUser)
    process {
        $g = (Get-AbStyle).Glyph
        $os = if ($Scan.Summary -and $Scan.Summary.PSObject.Properties['OS']) { [string]$Scan.Summary.OS } else { '' }
        $ms = (@($Scan.Results) | Measure-Object -Property DurationMs -Sum).Sum
        Write-Host ''
        Write-AbLine '  AutoByte health check' 'Accent+' "  $($g.Dot)  $($Scan.Computer)" Soft $(if ($os) { "  $($g.Dot)  $os" } else { '' }) Dim
        Write-AbRule
        Write-AbLine '  ' Text "$($g.Bad) $($Scan.Critical) critical" $(if ($Scan.Critical) { 'Bad+' } else { 'Dim' }) '     ' Text "$($g.Warn) $($Scan.Warning) warning(s)" $(if ($Scan.Warning) { 'Warn+' } else { 'Dim' }) '     ' Text "$($g.Info) $($Scan.Info) note(s)" Info "      $(@($Scan.Results).Count) checks in $(Format-AbDuration $ms)" Dim

        $groups = @(
            @{ Sev = 'Critical'; Title = 'Needs attention now'; Role = 'Bad'; Glyph = $g.Bad }
            @{ Sev = 'Warning'; Title = 'Worth fixing'; Role = 'Warn'; Glyph = $g.Warn }
            @{ Sev = 'Info'; Title = 'For your information'; Role = 'Info'; Glyph = $g.Info }
        )
        foreach ($grp in $groups) {
            $items = @($Scan.Findings | Where-Object Severity -EQ $grp.Sev)
            if (-not $items) { continue }
            Write-Host ''
            Write-AbLine '  ' Text $grp.Title ($grp.Role + '+')
            foreach ($f in $items) {
                Write-AbLine '   ' Text "$($grp.Glyph) " $grp.Role $f.Title Text
                if ($f.Detail) { Write-AbWrapped -Text $f.Detail -Indent 5 -Role Dim }
                if ($f.FixTask) {
                    $t = Get-AbTask -Id $f.FixTask
                    $how = if ($EndUser -and $t.EndUserSafe) { 'AutoByte can fix this (option 2)' } elseif ($EndUser) { 'Your IT support can fix this' } else { "abyte fix $($f.FixTask)" }
                    Write-AbLine '     ' Text "$($g.Arrow) " Key $how Key "  ($($t.Name))" Dim
                }
            }
        }
        $broken = @($Scan.Results | Where-Object { $_.Status -eq 'Failed' })
        $skipped = @($Scan.Results | Where-Object { $_.Status -eq 'Skipped' })
        if ($broken -or $skipped) {
            Write-Host ''
            foreach ($r in $broken) { Write-AbWrapped -Text "Could not run '$((Get-AbTask -Id $r.TaskId).Name)': $($r.Error)" -Indent 3 -Role Dim }
            if ($skipped) { Write-AbWrapped -Text ("Skipped (needs administrator): " + ((@($skipped | ForEach-Object { (Get-AbTask -Id $_.TaskId).Name })) -join ', ')) -Indent 3 -Role Dim }
        }
        Write-Host ''
        Write-AbRule
        if (-not @($Scan.Findings | Where-Object Severity -NE 'Info')) {
            if ($broken) { Write-AbLine '  ' Text "$($g.Ok) No problems found" 'Ok+' ", but $($broken.Count) check(s) could not run (see above)." Dim }
            else { Write-AbLine '  ' Text "$($g.Ok) Everything looks good." 'Ok+' }
        }
        elseif ($Scan.SuggestedFixes) {
            if ($EndUser) {
                $safe = @($Scan.SuggestedFixes | Where-Object { (Get-AbTask -Id $_).EndUserSafe })
                if ($safe) { Write-AbLine '  Next  ' Dim "choose 2 to fix $($safe.Count) safe item(s) automatically" Key }
                else { Write-AbLine '  Next  ' Dim 'contact your IT support (option 5 opens remote help)' Key }
            }
            else { Write-AbLine '  Next  ' Dim 'abyte fix --suggested' Key "   fixes $(@($Scan.SuggestedFixes).Count) item(s); add --whatif to preview" Dim }
        }
    }
}

function Write-AbFixView {
    <#
    .SYNOPSIS
        Shows fix results (Invoke-AbFix / Invoke-AbTask output) as a tidy, plain-language summary.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Results, [string]$Title = 'Results')
    $g = (Get-AbStyle).Glyph
    Write-Host ''
    Write-AbRule -Title "AutoByte $($g.Dot) $Title"
    $nameWidth = [Math]::Min(34, [Math]::Max(18, (@($Results | ForEach-Object { (Get-AbTask -Id $_.TaskId).Name.Length }) | Measure-Object -Maximum).Maximum + 2))
    foreach ($r in $Results) {
        $t = Get-AbTask -Id $r.TaskId
        $v = Get-AbStatusView -Status $r.Status -Kind $r.Kind
        Write-Host ''
        Write-AbLine '  ' Text "$($v.Glyph) " $v.Role $t.Name.PadRight($nameWidth) 'Text+' $v.Label ($v.Role + '+') "   $(Format-AbDuration $r.DurationMs)" Dim

        $data = if ($r.PSObject.Properties['Data'] -and $r.Data -is [System.Collections.IDictionary]) { $r.Data } else { @{} }
        $structured = $data.Contains('Updated') -or $data.Contains('Failed')
        foreach ($m in @($r.Messages)) {
            if ($structured -and $m -match '^(Updated |Not updated: |winget still lists |Fix needed: )') { continue }
            if ($m -match '^(Fixed and verified\.|Partly fixed\. Some items)') { continue }
            $text = $m -replace '^Fix needed: ', 'Found: ' -replace '^No action needed\.\s*', '' -replace '^Partly fixed\.\s*', '' -replace "^WhatIf: would run '(.+)'\.$", 'Preview: would run "$1". Nothing was changed.' -replace '^Fixed; a restart is required to complete\.$', 'Restart the PC to finish.'
            if ($text) { Write-AbWrapped -Text $text -Indent 5 -Role Dim }
        }
        if ($r.Error) { Write-AbWrapped -Text $r.Error -Indent 5 -Role Bad -FirstPrefix 'Error: ' }

        if ($structured) {
            $up = @($data['Updated']); $bad = @($data['Failed']); $still = @(if ($data.Contains('StillListed')) { $data['StillListed'] })
            $up = @($up | Where-Object { $_ }); $bad = @($bad | Where-Object { $_ }); $still = @($still | Where-Object { $_ })
            Write-AbLine '     ' Text "$($up.Count) updated" $(if ($up.Count) { 'Ok' } else { 'Dim' }) "  $($g.Dot)  " Dim "$($bad.Count) not updated" $(if ($bad.Count) { 'Bad' } else { 'Dim' }) $(if ($still.Count) { "  $($g.Dot)  $($still.Count) still listed" } else { '' }) Warn
            if ($up) {
                Write-AbLine '     Updated' 'Soft+'
                Write-AbWrapped -Text ((@($up | ForEach-Object { "$($_.Name) $($_.To)" })) -join ', ') -Indent 7 -Role Text
            }
            if ($bad) {
                Write-AbLine '     Not updated' 'Soft+'
                $w = [Math]::Min(26, (@($bad | ForEach-Object { ([string]$_.Name).Length }) | Measure-Object -Maximum).Maximum + 2)
                foreach ($b in $bad) {
                    Write-AbLine '       ' Text "$($g.Bad) " Bad ([string]$b.Name).PadRight($w) Text $b.Reason Text
                    if ($b.Hint) { Write-AbWrapped -Text $b.Hint -Indent (9 + $w) -Role Key -FirstPrefix "$($g.Arrow) " }
                }
            }
            if ($still) {
                Write-AbLine '     Still listed by winget ' 'Soft+' '(usually clears after a restart)' Dim
                Write-AbWrapped -Text ((@($still | ForEach-Object Name)) -join ', ') -Indent 7 -Role Text
            }
        }
    }
    Write-Host ''
    Write-AbRule
    $order = 'Fixed', 'Partial', 'Failed', 'NotNeeded', 'WhatIf', 'Skipped'
    $parts = @('  ', 'Text')
    foreach ($s in $order) {
        $n = @($Results | Where-Object Status -EQ $s).Count
        if (-not $n) { continue }
        $v = Get-AbStatusView -Status $s
        if ($parts.Count -gt 2) { $parts += @("  $($g.Dot)  ", 'Dim') }
        $parts += @("$n $($v.Label.ToLowerInvariant())", $v.Role)
    }
    $parts += @("      total $(Format-AbDuration ((@($Results) | Measure-Object -Property DurationMs -Sum).Sum))", 'Dim')
    Write-AbLine @parts
    if (@($Results | Where-Object RebootRequired).Count) { Write-AbLine '  Next  ' Dim 'Restart the PC to finish.' 'Key+' }
    $retry = @($Results | Where-Object { $_.Status -in 'Partial', 'Failed' } | ForEach-Object TaskId)
    if ($retry) { Write-AbLine '  Next  ' Dim "Follow the tips above, then run: abyte fix $($retry -join ' ')" Key }
}
