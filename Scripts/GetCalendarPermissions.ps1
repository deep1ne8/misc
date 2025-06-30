#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param(
    [string]$AdminUPN,
    [string]$ExportPath = "$env:USERPROFILE\Desktop\EXO_CalendarPermissions.csv",
    [ValidateSet("UserMailbox", "SharedMailbox", "RoomMailbox")]
    [string[]]$MailboxTypes = @("UserMailbox")
)

$script:logFile = "CalendarPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:logFile -Value $logEntry
    Write-Host $logEntry
}

function Connect-ToExchangeOnline {
    param([string]$UPN)
    try {
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress:$false -UseRPSSession:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online as $UPN" "SUCCESS"
    } catch {
        Write-Log "Failed to connect to Exchange Online: $_" "ERROR"
        exit 1
    }
}

function Export-CalendarPermissions {
    param([string]$Path)
    $results = @()
    try {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $MailboxTypes
        foreach ($mbx in $mailboxes) {
            $calendarPath = "$($mbx.UserPrincipalName):\Calendar"
            try {
                $permissions = Get-EXOMailboxFolderPermission -Identity $calendarPath -ErrorAction Stop
                foreach ($perm in $permissions) {
                    $results += [pscustomobject]@{
                        Mailbox = $mbx.UserPrincipalName
                        Folder = "Calendar"
                        User = $perm.User.DisplayName
                        AccessRights = $perm.AccessRights -join ", "
                    }
                }
            } catch {
                Write-Log "Cannot get permissions for $calendarPath: $_" "WARNING"
            }
        }
        if ($results.Count -gt 0) {
            $results | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            Write-Log "Export complete: $Path" "SUCCESS"
        } else {
            Write-Log "No permissions found to export." "WARNING"
        }
    } catch {
        Write-Log "Error during mailbox enumeration or permission export: $_" "ERROR"
    }
}

# Main
Write-Log "=== Calendar Permissions Export Started ==="
if (-not $AdminUPN) {
    $AdminUPN = Read-Host "Enter your Exchange Online Admin UPN"
}
Connect-ToExchangeOnline -UPN $AdminUPN
Export-CalendarPermissions -Path $ExportPath
Write-Log "=== Calendar Permissions Export Completed ==="
