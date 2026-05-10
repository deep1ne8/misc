#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Tracks file deletions from a user's Desktop on a remote server using Security logs and Recycle Bin.

.DESCRIPTION
    Queries Security Event 4663 (object access) and Recycle Bin to identify files deleted from
    a specific user's Desktop within a given timeframe. Requires appropriate audit policies enabled.

.NOTES
    Prerequisites:
    - Remote server must have File Audit policies enabled (Success/Failure for Delete operations)
    - WinRM enabled on target server
    - Appropriate permissions to read Security logs remotely
#>

[CmdletBinding()]
param()

# ============================================
# PARAMETER COLLECTION
# ============================================

function Get-ValidatedComputerName {
    do {
        $computer = Read-Host -Prompt "Enter remote server name or IP"
        if ([string]::IsNullOrWhiteSpace($computer)) {
            Write-Warning "Computer name cannot be empty."
            continue
        }
        
        Write-Host "Testing connection to $computer..." -ForegroundColor Cyan
        if (Test-Connection -ComputerName $computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            Write-Host "✓ Connection successful" -ForegroundColor Green
            return $computer
        } else {
            Write-Warning "Cannot reach $computer. Please verify the name/IP and network connectivity."
            $retry = Read-Host "Try again? (Y/N)"
            if ($retry -notmatch '^y') { throw "Operation cancelled by user." }
        }
    } while ($true)
}

function Get-ValidatedUsername {
    do {
        $user = Read-Host -Prompt "Enter local username (folder name under C:\Users\)"
        if ([string]::IsNullOrWhiteSpace($user)) {
            Write-Warning "Username cannot be empty."
            continue
        }
        return $user.Trim()
    } while ($true)
}

function Get-ValidatedDateRange {
    do {
        $startInput = Read-Host -Prompt "Enter START date/time (e.g., 2025-09-24 or 2025-09-24T00:00:00)"
        try {
            $start = Get-Date $startInput -ErrorAction Stop
        } catch {
            Write-Warning "Invalid date format. Please use YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS"
            continue
        }

        $endInput = Read-Host -Prompt "Enter END date/time (e.g., 2025-09-30T23:59:59)"
        try {
            $end = Get-Date $endInput -ErrorAction Stop
        } catch {
            Write-Warning "Invalid date format. Please use YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS"
            continue
        }

        if ($end -le $start) {
            Write-Warning "End date must be after start date."
            continue
        }

        $span = $end - $start
        Write-Host "Date range: $($span.Days) days, $($span.Hours) hours" -ForegroundColor Cyan
        return @{ Start = $start; End = $end }
    } while ($true)
}

function Get-OutputPath {
    $defaultName = "Desktop_Deletes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $path = Read-Host -Prompt "Enter output CSV path (press Enter for default: $PWD\$defaultName)"
    
    if ([string]::IsNullOrWhiteSpace($path)) {
        return Join-Path $PWD $defaultName
    }
    
    # Ensure .csv extension
    if ($path -notmatch '\.csv$') {
        $path += '.csv'
    }
    
    return $path
}

# ============================================
# COLLECT PARAMETERS
# ============================================

Write-Host "`n=== Remote Desktop Delete Tracker ===" -ForegroundColor Yellow
Write-Host "This script tracks file deletions from a user's Desktop folder.`n"

$ComputerName = Get-ValidatedComputerName
$Username     = Get-ValidatedUsername
$dateRange    = Get-ValidatedDateRange
$Start        = $dateRange.Start
$End          = $dateRange.End
$OutCsv       = Get-OutputPath

Write-Host "`n--- Configuration Summary ---" -ForegroundColor Yellow
Write-Host "Remote Server : $ComputerName"
Write-Host "Username      : $Username"
Write-Host "Start Date    : $Start"
Write-Host "End Date      : $End"
Write-Host "Output File   : $OutCsv"

$confirm = Read-Host "`nProceed with these settings? (Y/N)"
if ($confirm -notmatch '^y') {
    Write-Host "Operation cancelled." -ForegroundColor Red
    exit
}

# ============================================
# REMOTE EXECUTION
# ============================================

Write-Host "`nConnecting to $ComputerName..." -ForegroundColor Cyan

try {
    $results = Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
        param($Username, $Start, $End)

        # Resolve Desktop path and SID from ProfileList
        $profileSid = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue |
            Where-Object {
                $profilePath = (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue).ProfileImagePath
                $profilePath -ieq "C:\Users\$Username"
            } | Select-Object -First 1 -ExpandProperty PSChildName

        $desktopPath = "C:\Users\$Username\Desktop"

        if (-not (Test-Path $desktopPath)) {
            throw "Desktop path does not exist: $desktopPath"
        }

        # --- 1) Security Log: File delete operations (Event 4663) ---
        $secResults = @()
        try {
            $filter = @{
                LogName   = 'Security'
                Id        = 4663  # Object access event
                StartTime = $Start
                EndTime   = $End
            }
            
            $events = @(Get-WinEvent -FilterHashtable $filter -ErrorAction Stop)
            
            foreach ($e in $events) {
                $xml = [xml]$e.ToXml()
                $eventData = @{}
                
                foreach ($data in $xml.Event.EventData.Data) {
                    $eventData[$data.Name] = $data.'#text'
                }

                $objectName = $eventData['ObjectName']
                $accessList = $eventData['AccessList']
                $subjectUser = $eventData['SubjectUserName']
                $subjectDomain = $eventData['SubjectDomainName']

                # Filter: Desktop files with DELETE operations
                if ($objectName -and ($objectName -like "$desktopPath*")) {
                    if ($accessList -match 'DELETE|DELETE_CHILD|WriteData') {
                        $secResults += [PSCustomObject]@{
                            Source      = 'SecurityLog'
                            TimeCreated = $e.TimeCreated
                            User        = "$subjectDomain\$subjectUser"
                            Path        = $objectName
                            Action      = 'DELETE (4663)'
                            Details     = $accessList
                        }
                    }
                }
            }
        } catch [System.Exception] {
            if ($_.Exception.Message -match 'No events were found') {
                # This is normal if no matching events exist
                $secResults += [PSCustomObject]@{
                    Source='SecurityLog'; TimeCreated=$null; User=$null; Path=$null; 
                    Action='INFO'; Details='No security events found for the specified time range'
                }
            } else {
                $secResults += [PSCustomObject]@{
                    Source='SecurityLog'; TimeCreated=$null; User=$null; Path=$null; 
                    Action='ERROR'; Details=$_.Exception.Message
                }
            }
        }

        # --- 2) Recycle Bin: Deleted items analysis ---
        $rbResults = @()
        $shell = $null
        try {
            $shell = New-Object -ComObject Shell.Application
            $recycleBin = $shell.Namespace(0xA)
            
            if ($recycleBin) {
                $items = @($recycleBin.Items())
                
                foreach ($item in $items) {
                    $itemName = $recycleBin.GetDetailsOf($item, 0)
                    $origLocation = $recycleBin.GetDetailsOf($item, 1)
                    $dateDeleted = $recycleBin.GetDetailsOf($item, 2)

                    $deletedDate = $null
                    if ([DateTime]::TryParse($dateDeleted, [ref]$deletedDate)) {
                        if ($deletedDate -ge $Start -and $deletedDate -le $End) {
                            $fullPath = if ($origLocation) { 
                                Join-Path $origLocation $itemName 
                            } else { 
                                $itemName 
                            }

                            if ($fullPath -like "$desktopPath*") {
                                $rbResults += [PSCustomObject]@{
                                    Source      = 'RecycleBin'
                                    TimeCreated = $deletedDate
                                    User        = $Username
                                    Path        = $fullPath
                                    Action      = 'Deleted → Recycle Bin'
                                    Details     = 'Still in Recycle Bin'
                                }
                            }
                        }
                    }
                }
            }
        } catch {
            $rbResults += [PSCustomObject]@{
                Source='RecycleBin'; TimeCreated=$null; User=$null; Path=$null; 
                Action='ERROR'; Details=$_.Exception.Message
            }
        } finally {
            if ($shell) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
        }

        # --- 3) Audit log tampering detection (Event 1102) ---
        $tamperEvents = @()
        try {
            $tamperEvents = @(Get-WinEvent -FilterHashtable @{
                LogName='Security'; Id=1102; StartTime=$Start; EndTime=$End
            } -ErrorAction SilentlyContinue)
        } catch { }

        [PSCustomObject]@{
            DesktopPath = $desktopPath
            ProfileSID  = $profileSid
            SecurityLog = $secResults
            RecycleBin  = $rbResults
            TamperCount = $tamperEvents.Count
            TamperEvents = $tamperEvents | Select-Object TimeCreated, Message
        }

    } -ArgumentList $Username, $Start, $End

    # ============================================
    # PROCESS RESULTS
    # ============================================

    Write-Host "`n✓ Data collection complete" -ForegroundColor Green
    Write-Host "Desktop Path: $($results.DesktopPath)"
    Write-Host "Profile SID : $($results.ProfileSID)"

    # Combine all results
    $allFindings = @($results.SecurityLog) + @($results.RecycleBin)
    
    if ($allFindings.Count -eq 0) {
        Write-Host "`nNo deletion events found for the specified criteria." -ForegroundColor Yellow
    } else {
        Write-Host "`nFound $($allFindings.Count) deletion event(s):" -ForegroundColor Cyan
        $allFindings | Sort-Object TimeCreated | Format-Table -AutoSize
        
        # Export to CSV
        $allFindings | Sort-Object TimeCreated | 
            Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutCsv -Force
        
        Write-Host "`n✓ Results saved to: $OutCsv" -ForegroundColor Green
    }

    # Check for audit log tampering
    if ($results.TamperCount -gt 0) {
        Write-Warning "`n⚠ SECURITY ALERT: Audit log was cleared $($results.TamperCount) time(s) during the specified window!"
        Write-Warning "This may indicate tampering. Review these events:"
        $results.TamperEvents | Format-Table -AutoSize
    }

} catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
    Write-Host "`n✗ ERROR: Cannot connect to $ComputerName" -ForegroundColor Red
    Write-Host "Possible causes:" -ForegroundColor Yellow
    Write-Host "  - WinRM is not enabled on the remote server"
    Write-Host "  - Firewall is blocking WinRM (port 5985/5986)"
    Write-Host "  - You don't have permission to connect"
    Write-Host "`nTo enable WinRM on the remote server, run: Enable-PSRemoting -Force"
    exit 1
} catch {
    Write-Host "`n✗ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}

Write-Host "`nScript completed successfully." -ForegroundColor Green
