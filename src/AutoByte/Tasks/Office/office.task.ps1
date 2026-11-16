@{
    Id          = 'office'
    Name        = 'Microsoft 365 Apps'
    Kind        = 'Check'
    Category    = 'Office'
    Description = 'Reports the installed Click-to-Run Office version and channel, and whether automatic updates are turned off.'
    Test        = {
        param($Ctx, $P)
        $key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $version = Get-AbRegistryValue -Path $key -Name 'VersionToReport'
        if (-not $version) { return }
        $platform = Get-AbRegistryValue -Path $key -Name 'Platform'
        $channel = Get-AbRegistryValue -Path $key -Name 'UpdateChannel'
        if (-not $channel) { $channel = Get-AbRegistryValue -Path $key -Name 'CDNBaseUrl' }
        $known = @{ '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'; '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise'; '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise'; '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current (Preview)' }
        $chanName = 'Unknown'
        foreach ($g in $known.Keys) { if ("$channel" -match $g) { $chanName = $known[$g] } }
        New-AbFinding -Code 'office.info' -Severity Info -Title "Microsoft 365 Apps $version ($platform, $chanName channel)" -Data @{ Version = "$version"; Platform = "$platform"; Channel = $chanName }
        if ("$(Get-AbRegistryValue -Path $key -Name 'UpdatesEnabled')" -eq 'False') {
            New-AbFinding -Code 'office.updates-off' -Severity Warning -Title 'Office automatic updates are turned off' -Detail 'Office will miss security fixes unless updates are managed another way (Intune, WSUS, ConfigMgr).'
        }
    }
}
