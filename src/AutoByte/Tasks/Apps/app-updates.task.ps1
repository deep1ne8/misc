@{
    Id          = 'app-updates'
    Name        = 'Outdated applications'
    Kind        = 'Check'
    Category    = 'Apps'
    Description = 'Lists installed applications with a newer version available (winget). Out-of-date browsers, PDF readers, runtimes and archivers are flagged as a security risk.'
    Test        = {
        param($Ctx, $P)
        $apps = Get-AbAppUpgrade
        if ($null -eq $apps) { return New-AbFinding -Code 'apps.winget-missing' -Severity Info -Title 'winget is not available' -Detail 'Install App Installer from the Microsoft Store (or abyte legacy Install-Winget) to check application updates.' }
        $apps = @($apps)
        if (-not $apps) { return }
        $risky = '(?i)chrome|edge|firefox|brave|opera|acrobat|reader|foxit|java|jre|jdk|zoom|teams|7-?zip|winrar|notepad\+\+|vlc|putty|filezilla|python|node|dotnet|\.net|vcredist|anydesk|teamviewer'
        $hot = @($apps | Where-Object { "$($_.Name) $($_.Id)" -match $risky })
        $sev = if ($hot) { 'Warning' } else { 'Info' }
        $show = @(($hot + $apps) | Select-Object -Unique -First 8 | ForEach-Object { "$($_.Name) $($_.Version) -> $($_.Available)" })
        New-AbFinding -Code 'apps.updates-available' -Severity $sev -Title "$($apps.Count) app update(s) available, $($hot.Count) security-sensitive" -Detail ($show -join '; ') -FixTask 'update-apps' `
            -Data @{ Count = $apps.Count; Ids = (@($apps | ForEach-Object Id) -join ',') }
    }
}
