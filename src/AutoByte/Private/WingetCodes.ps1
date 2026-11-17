# winget / installer exit codes -> plain-language reason, what to do, and outcome class.
# Source: microsoft/winget-cli doc/windows/package-manager/winget/returnCodes.md
# Outcome: Ok (treat as updated), Reboot (updated, restart needed), Current (already up to date), Fail.

$script:AbWingetCodes = @{
    0x8A150101 = @('Fail', 'The app was open.', 'Close it, then run the update again.')
    0x8A150103 = @('Fail', 'A file of the app was in use.', 'Close the app (or restart), then run the update again.')
    0x8A150111 = @('Fail', 'Another program was using the app.', 'Restart the PC, then run the update again.')
    0x8A150102 = @('Fail', 'Another installation was already running.', 'Wait a few minutes and try again.')
    0x8A15007D = @('Fail', 'It is installed for one user only, so an administrator session cannot update it.', 'Update it while signed in as that user (run abyte without "Run as administrator"), or from the app itself.')
    0x8A150056 = @('Fail', 'Its installer refuses to run as administrator.', 'Update it while signed in as the user, without "Run as administrator".')
    0x8A150030 = @('Fail', "The old version's uninstaller failed.", 'Uninstall it in Settings > Apps, then reinstall the latest version.')
    0x8A150114 = @('Fail', 'Its installer cannot upgrade an existing copy.', 'Uninstall it in Settings > Apps, then reinstall the latest version.')
    0x8A150104 = @('Fail', 'A component it needs is missing.', 'Run Windows Update, then try again.')
    0x8A150110 = @('Fail', 'A component it needs failed to install.', 'Run Windows Update, then try again.')
    0x8A150105 = @('Fail', 'The disk is full.', "Free up space (abyte fix clean-temp-files), then try again.")
    0x8A150106 = @('Fail', 'Not enough memory.', 'Close other apps, then try again.')
    0x8A150107 = @('Fail', 'No internet connection.', 'Connect to the internet, then try again.')
    0x8A15010A = @('Fail', 'A restart is needed before it can install.', 'Restart the PC, then run the update again.')
    0x8A15010C = @('Fail', 'The installation was cancelled.', 'Run the update again and let it finish.')
    0x8A15010F = @('Fail', 'Blocked by your organisation''s policy.', 'Ask your IT support.')
    0x8A15003A = @('Fail', 'Blocked by Group Policy.', 'Ask your IT support.')
    0x8A15001B = @('Fail', 'The Microsoft Store is blocked by policy.', 'Ask your IT support.')
    0x8A15001C = @('Fail', 'This Store app is blocked by policy.', 'Ask your IT support.')
    0x8A150011 = @('Fail', 'The download did not match the expected file.', 'Usually the vendor just released a new build; try again tomorrow.')
    0x8A150068 = @('Fail', 'It is pinned to its current version.', 'Unpin it (winget pin remove) if you want it updated.')
    0x8A150108 = @('Fail', 'Its installer reported an error.', 'Contact the vendor or your IT support.')
    0x8A150049 = @('Fail', 'The MSI installer failed.', 'Restart, then try again; if it keeps failing, reinstall the app.')
    0x8A150115 = @('Fail', 'Its installer reported a custom error.', 'Reinstall the app or contact the vendor.')
    0x8A150113 = @('Fail', 'This version does not support this PC.', 'Nothing to do; keep the current version.')
    0x8A150010 = @('Fail', 'No installer fits this PC.', 'Nothing to do; keep the current version.')
    0x8A150109 = @('Reboot', 'Updated; a restart finishes it.', '')
    0x8A15010B = @('Reboot', 'Updated; the PC will restart to finish.', '')
    0x8A15002B = @('Current', 'Already up to date.', '')
    0x8A15004F = @('Current', 'Already up to date.', '')
    0x8A15010D = @('Current', 'Another version is already installed.', '')
    0x8A15010E = @('Current', 'A newer version is already installed.', '')
    0x8A150061 = @('Current', 'Already installed.', '')
}

function Get-AbWingetOutcome {
    # Returns Outcome (Ok|Reboot|Current|Fail), Code (0x........), Reason, Hint for a winget exit code.
    param([Parameter(Mandatory)][long]$ExitCode)
    $hex = '0x{0:X8}' -f ([uint32]($ExitCode -band 0xFFFFFFFFL))
    if ($ExitCode -eq 0) { return [pscustomobject]@{ Outcome = 'Ok'; Code = $hex; Reason = ''; Hint = '' } }
    if ($ExitCode -in 3010, 1641) { return [pscustomobject]@{ Outcome = 'Reboot'; Code = $hex; Reason = 'Updated; a restart finishes it.'; Hint = '' } }
    if ($ExitCode -eq 1603) { return [pscustomobject]@{ Outcome = 'Fail'; Code = $hex; Reason = 'The installer failed (MSI error 1603).'; Hint = 'Restart, then try again; if it keeps failing, reinstall the app.' } }
    $key = [int64]([uint32]($ExitCode -band 0xFFFFFFFFL))
    foreach ($k in $script:AbWingetCodes.Keys) {
        if (([int64]$k -band 0xFFFFFFFFL) -eq $key) {
            $v = $script:AbWingetCodes[$k]
            return [pscustomobject]@{ Outcome = $v[0]; Code = $hex; Reason = $v[1]; Hint = $v[2] }
        }
    }
    [pscustomobject]@{ Outcome = 'Fail'; Code = $hex; Reason = "The installer failed ($hex)."; Hint = 'Run "winget upgrade --id <id>" to see the details.' }
}
