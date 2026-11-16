@{
    Id          = 'temp-files'
    Name        = 'Temporary files'
    Kind        = 'Check'
    Category    = 'Storage'
    Description = 'Space used by old files in the Windows and user temp folders.'
    Test        = {
        param($Ctx, $P)
        $cutoff = (Get-AbDate).AddDays(-[int]$Ctx.Config.TempFileMinAgeDays)
        $bytes = 0
        foreach ($folder in @((Get-AbSystemPath -Name UserTemp), (Get-AbSystemPath -Name WindowsTemp)) | Sort-Object -Unique) {
            foreach ($f in @(Get-AbChildItem -Path $folder -Recurse)) { if ($f.LastWriteTime -lt $cutoff) { $bytes += $f.Length } }
        }
        $gb = [math]::Round($bytes / 1GB, 2)
        if ($gb -ge 2) {
            New-AbFinding -Code 'temp.large' -Severity Warning -Title "$gb GB of old temporary files" -Detail 'Safe to remove.' -FixTask 'clean-temp-files' -Data @{ GB = $gb }
        }
        elseif ($gb -ge 0.5) {
            New-AbFinding -Code 'temp.some' -Severity Info -Title "$gb GB of old temporary files" -FixTask 'clean-temp-files' -Data @{ GB = $gb }
        }
    }
}
