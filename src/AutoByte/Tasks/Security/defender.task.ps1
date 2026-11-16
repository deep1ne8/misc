@{
    Id          = 'defender'
    Name        = 'Microsoft Defender'
    Kind        = 'Check'
    Category    = 'Security'
    Description = 'Real-time protection and virus definition age (skipped when another antivirus is in charge).'
    Test        = {
        param($Ctx, $P)
        try { $s = Get-AbDefenderStatus }
        catch { return (New-AbFinding -Code 'defender.unavailable' -Severity Info -Title 'Microsoft Defender status unavailable' -Detail 'Another antivirus may be installed, or Defender is disabled by policy.') }
        if (-not $s.AntivirusEnabled) {
            return (New-AbFinding -Code 'defender.passive' -Severity Info -Title 'Defender antivirus is not the active antivirus' -Detail 'Expected if another antivirus product is installed.')
        }
        if (-not $s.RealTimeProtectionEnabled) {
            New-AbFinding -Code 'defender.realtime-off' -Severity Critical -Title 'Real-time virus protection is off' -Detail 'Turn it back on in Windows Security, or check whether a policy disabled it.'
        }
        if ([int]$s.AntivirusSignatureAge -gt [int]$Ctx.Config.DefenderSigMaxDays) {
            New-AbFinding -Code 'defender.signatures-old' -Severity Warning -Title "Virus definitions are $($s.AntivirusSignatureAge) days old" -FixTask 'update-defender-signatures' -Data @{ Days = [int]$s.AntivirusSignatureAge }
        }
    }
}
