#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Modern Exchange Online Calendar Permissions Management Tool (EXO V3)
.DESCRIPTION
    Uses latest Get-EXO* cmdlets for optimal performance and modern Exchange Online management:
    - Mode 1: List and export all active mailboxes using Get-EXOMailbox
    - Mode 2: Import CSV and apply calendar permissions with safety checks
.PARAMETER Mode
    Operation mode: List, Apply, or Interactive
.PARAMETER AdminUPN
    Admin User Principal Name for Exchange Online connection
.PARAMETER Permission
    Calendar permission level to apply (AvailabilityOnly, LimitedDetails, Reviewer, Editor)
.PARAMETER ImportPath
    Path to CSV file containing mailboxes to process
.PARAMETER ExportPath
    Path for exporting mailbox list CSV
.PARAMETER WhatIf
    Preview changes without applying them
#>

[CmdletBinding()]
param(
    [ValidateSet("List", "Apply", "Interactive")]
    [string]$Mode = "Interactive",
    [string]$AdminUPN,
    [string]$Permission,
    [string]$ImportPath,
    [string]$ExportPath,
    [string[]]$PermissionFilter = @("Reviewer"),
    [ValidateSet("UserMailbox", "SharedMailbox", "RoomMailbox")]
    [string[]]$MailboxTypes = @("UserMailbox"),
    [switch]$WhatIf
)

$script:logFile = "ModernEXOCalendarTool_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:validPermissions = @("AvailabilityOnly", "LimitedDetails", "Reviewer", "Editor")

function Write-Log {
    param(
        [string]$Message, 
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "VERBOSE")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:logFile -Value $logEntry -ErrorAction SilentlyContinue
    switch ($Level) {
        "ERROR" { Write-Host "❌ $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "⚠️  $Message" -ForegroundColor Yellow }
        "SUCCESS" { Write-Host "✅ $Message" -ForegroundColor Green }
        "VERBOSE" { Write-Host "ℹ️  $Message" -ForegroundColor Cyan }
        default { Write-Host "📝 $Message" -ForegroundColor White }
    }
}

function Write-Progress-Enhanced {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total,
        [string]$CurrentItem = ""
    )
    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    $statusText = "$Status - $Current of $Total ($percent%)"
    if ($CurrentItem) { $statusText += " - $CurrentItem" }
    Write-Progress -Activity $Activity -Status $statusText -PercentComplete $percent
    Write-Log "Progress: ${statusText}" "VERBOSE"
}

function Test-ModernExchangeConnection {
    try { $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop; return $true } catch { return $false }
}

function Connect-ToModernExchangeOnline {
    param([string]$UPN)
    if (-not (Test-ModernExchangeConnection)) {
        if (-not $UPN) { $UPN = Read-Host "Enter your admin UPN for Exchange Online" }
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress:$false -UseRPSSession:$false -ErrorAction Stop
    }
    return $true
}

function Export-CalendarPermissionsForAllMailboxes {
    param ([string]$OutputPath = "$env:USERPROFILE\Desktop\EXO_CalendarPermissions.csv")
    $results = @()
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $MailboxTypes
    $total = $mailboxes.Count
    $i = 0
    foreach ($mbx in $mailboxes) {
        $i++
        Write-Progress-Enhanced -Activity "Exporting Permissions" -Status "Checking" -Current $i -Total $total -CurrentItem $mbx.UserPrincipalName
        try {
            $calendarPath = "$($mbx.UserPrincipalName):\Calendar"
            $perms = Get-EXOMailboxFolderPermission -Identity $calendarPath -ErrorAction Stop
            foreach ($perm in $perms) {
                if ($PermissionFilter -contains $perm.AccessRights.ToString()) {
                    $results += [pscustomobject]@{
                        Mailbox = $mbx.UserPrincipalName
                        Folder = "Calendar"
                        User = $perm.User.DisplayName
                        AccessRights = $perm.AccessRights -join ", "
                        SharingPermissionFlags = $perm.SharingPermissionFlags
                    }
                }
            }
        } catch { Write-Log "Failed for $($mbx.UserPrincipalName): $_" "WARNING" }
    }
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Log "Exported to $OutputPath" "SUCCESS"
    } else {
        Write-Log "No matching permission data found." "WARNING"
    }
}

function Import-MailboxesFromCSV {
    param([string]$Path)
    if (-not (Test-Path $Path)) { Write-Log "CSV not found: ${Path}" "ERROR"; return @() }
    $csv = Import-Csv -Path $Path
    $col = ($csv | Get-Member -MemberType NoteProperty).Name | Where-Object { $_ -match "Email|UPN" } | Select-Object -First 1
    if (-not $col) { Write-Log "No valid email/UPN column in CSV." "ERROR"; return @() }
    return $csv | ForEach-Object { $_.$col }
}

function Set-ModernCalendarPermissions {
    param([string[]]$EmailAddresses, [string]$Permission, [switch]$WhatIf)
    foreach ($email in $EmailAddresses) {
        $path = "${email}:\Calendar"
        try {
            $existing = Get-EXOMailboxFolderPermission -Identity $path -User Default -ErrorAction SilentlyContinue
            if ($existing -and $existing.AccessRights -contains $Permission) {
                Write-Log "$email already has $Permission" "INFO"
                continue
            }
            if ($WhatIf) {
                Write-Log "[WHATIF] Would apply $Permission to $email" "INFO"
            } else {
                if ($existing) {
                    Set-MailboxFolderPermission -Identity $path -User Default -AccessRights $Permission -Confirm:$false
                    Write-Log "Updated $email to $Permission" "SUCCESS"
                } else {
                    Add-MailboxFolderPermission -Identity $path -User Default -AccessRights $Permission -Confirm:$false
                    Write-Log "Added $Permission to $email" "SUCCESS"
                }
            }
        } catch { Write-Log "Failed for ${email}: $_" "ERROR" }
    }
}

function Show-InteractiveMenu {
    Write-Host "\n=== Modern EXO Calendar Tool ===" -ForegroundColor Cyan
    Write-Host "1. List and Export All Mailboxes with Filtered Calendar Permissions"
    Write-Host "2. Import CSV and Apply Permissions"
    Write-Host "3. Preview Permission Changes"
    Write-Host "4. Test EXO Connection"
    Write-Host "5. Exit"
}

try {
    Write-Log "Tool Started" "INFO"
    Connect-ToModernExchangeOnline -UPN $AdminUPN | Out-Null
    switch ($Mode) {
        "List" {
            Export-CalendarPermissionsForAllMailboxes -OutputPath $ExportPath
        }
        "Apply" {
            $emails = Import-MailboxesFromCSV -Path $ImportPath
            if ($emails.Count -gt 0) {
                Set-ModernCalendarPermissions -EmailAddresses $emails -Permission $Permission -WhatIf:$WhatIf
            }
        }
        "Interactive" {
            do {
                Show-InteractiveMenu
                $sel = Read-Host "Select option"
                switch ($sel) {
                    "1" {
                        $ExportPath = Read-Host "Enter export path or press Enter for default"
                        if ([string]::IsNullOrWhiteSpace($ExportPath)) { $ExportPath = "$env:USERPROFILE\Desktop\EXO_CalendarPermissions.csv" }
                        Export-CalendarPermissionsForAllMailboxes -OutputPath $ExportPath
                    }
                    "2" {
                        $ImportPath = Read-Host "CSV path to import"
                        $Permission = Read-Host "Enter permission to apply"
                        $emails = Import-MailboxesFromCSV -Path $ImportPath
                        Set-ModernCalendarPermissions -EmailAddresses $emails -Permission $Permission
                    }
                    "3" {
                        $ImportPath = Read-Host "CSV path to import"
                        $Permission = Read-Host "Enter permission to preview"
                        $emails = Import-MailboxesFromCSV -Path $ImportPath
                        Set-ModernCalendarPermissions -EmailAddresses $emails -Permission $Permission -WhatIf
                    }
                    "4" {
                        if (Test-ModernExchangeConnection) { Write-Log "✅ Connected" "SUCCESS" } else { Write-Log "❌ Not Connected" "ERROR" }
                    }
                }
            } while ($sel -ne "5")
        }
    }
    Write-Log "Tool Finished" "SUCCESS"
} catch {
    Write-Log "Failed: $($_.Exception.Message)" "ERROR"
} finally {
    Write-Log "=== End ===" "INFO"
}

