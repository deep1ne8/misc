@{
    Id          = 'pending-reboot'
    Name        = 'Pending restart'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Detects a restart that Windows Update, servicing or file operations are waiting for.'
    Test        = {
        param($Ctx, $P)
        $reasons = @()
        if (Test-AbPath -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'Windows servicing' }
        if (Test-AbPath -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons += 'Windows Update' }
        if (Get-AbRegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations') { $reasons += 'pending file operations' }
        if ($reasons) {
            New-AbFinding -Code 'system.reboot-pending' -Severity Warning -Title 'A restart is pending' -Detail ("Requested by: " + ($reasons -join ', ') + '. Restart when convenient so updates finish installing.') -Data @{ Reasons = $reasons }
        }
    }
}
