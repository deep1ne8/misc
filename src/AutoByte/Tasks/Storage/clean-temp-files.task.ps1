@{
    Id          = 'clean-temp-files'
    Name        = 'Clean temporary files'
    Kind        = 'Fix'
    Category    = 'Storage'
    Description = 'Deletes files older than MinAgeDays from the user temp folder (and Windows\Temp when elevated). Files in use are skipped.'
    EndUserSafe = $true
    Parameters  = @{ MinAgeDays = 7; MinReclaimMB = 100 }
    Test        = {
        param($Ctx, $P)
        $cutoff = (Get-AbDate).AddDays(-[int]$P.MinAgeDays)
        $folders = @(Get-AbSystemPath -Name UserTemp)
        if (Test-AbIsAdmin) { $folders += Get-AbSystemPath -Name WindowsTemp }
        $files = @(foreach ($folder in ($folders | Sort-Object -Unique)) { Get-AbChildItem -Path $folder -Recurse | Where-Object { $_.LastWriteTime -lt $cutoff } })
        $mb = [math]::Round((($files | Measure-Object -Property Length -Sum).Sum) / 1MB, 1)
        if ($Ctx.AfterRepair) {
            # What is left is in use by running programs; that is expected, not a failure.
            return @{ Needed = $false; Reason = "$mb MB still present (files in use)." }
        }
        $Ctx.Candidates = $files
        $mb = [math]::Round((($files | Measure-Object -Property Length -Sum).Sum) / 1MB, 1)
        if ($mb -ge [double]$P.MinReclaimMB) { @{ Needed = $true; Reason = "$mb MB in $($files.Count) old temp files." } }
        else { @{ Needed = $false; Reason = "Only $mb MB of old temp files (threshold $($P.MinReclaimMB) MB)." } }
    }
    Repair      = {
        param($Ctx, $P)
        $removed = 0; $skipped = 0; $bytes = 0
        foreach ($f in @($Ctx.Candidates)) {
            if (Remove-AbFile -Path $f.FullName) { $removed++; $bytes += $f.Length } else { $skipped++ }
        }
        Add-AbMessage -Result $Ctx.Result -Message ("Removed {0} files ({1} MB); skipped {2} in use." -f $removed, [math]::Round($bytes / 1MB, 1), $skipped)
        @{ RebootRequired = $false }
    }
}
