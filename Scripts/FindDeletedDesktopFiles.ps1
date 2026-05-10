#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Forensic analysis to find files already deleted from Desktop.

.DESCRIPTION
    Checks multiple sources to find deleted files:
    1. Recycle Bin (still recoverable)
    2. USN Journal (if retention period hasn't expired)
    3. Shadow Copies/Previous Versions (if enabled)
    4. Security Event Log (if auditing was enabled)
    5. PowerShell History (if deletion was via PowerShell)

.NOTES
    Run this ASAP after discovering deletions - USN journal has limited retention.
#>

[CmdletBinding()]
param()

# ============================================
# PARAMETER COLLECTION
# ============================================

Write-Host "`n=== Deleted File Forensic Analysis ===" -ForegroundColor Yellow
Write-Host "This will check all available sources for deleted files.`n"

$computer = Read-Host "Enter server name (or press Enter for local machine)"
if ([string]::IsNullOrWhiteSpace($computer)) { 
    $computer = $env:COMPUTERNAME 
    $isLocal = $true
} else {
    $isLocal = $false
}

$username = Read-Host "Enter username (folder name under C:\Users\)"

$startInput = Read-Host "Enter approximate START date of deletions (YYYY-MM-DD or YYYY-MM-DD HH:MM)"
$start = Get-Date $startInput

$endInput = Read-Host "Enter approximate END date (or press Enter for now)"
if ([string]::IsNullOrWhiteSpace($endInput)) {
    $end = Get-Date
} else {
    $end = Get-Date $endInput
}

Write-Host "`nAnalyzing period: $start to $end" -ForegroundColor Cyan
Write-Host "Target: $computer\$username\Desktop`n" -ForegroundColor Cyan

$outputFile = Join-Path $PWD "DeletedFiles_$($username)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# ============================================
# FORENSIC ANALYSIS SCRIPT BLOCK
# ============================================

$scriptBlock = {
    param($Username, $Start, $End)
    
    $desktopPath = "C:\Users\$Username\Desktop"
    $allFindings = @()
    
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "FORENSIC ANALYSIS REPORT" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Desktop Path: $desktopPath"
    Write-Host "Time Window: $Start to $End"
    Write-Host ""
    
    # ============================================
    # SOURCE 1: RECYCLE BIN (Highest Priority)
    # ============================================
    Write-Host "[1/5] Checking Recycle Bin..." -ForegroundColor Cyan
    $rbCount = 0
    
    try {
        $shell = New-Object -ComObject Shell.Application
        $recycleBin = $shell.Namespace(0xA)
        
        if ($recycleBin) {
            $items = @($recycleBin.Items())
            
            foreach ($item in $items) {
                $itemName = $recycleBin.GetDetailsOf($item, 0)
                $origLocation = $recycleBin.GetDetailsOf($item, 1)
                $dateDeleted = $recycleBin.GetDetailsOf($item, 2)
                $itemSize = $recycleBin.GetDetailsOf($item, 3)
                
                $deletedDate = $null
                if ([DateTime]::TryParse($dateDeleted, [ref]$deletedDate)) {
                    if ($deletedDate -ge $Start -and $deletedDate -le $End) {
                        $fullPath = if ($origLocation) { 
                            Join-Path $origLocation $itemName 
                        } else { 
                            $itemName 
                        }
                        
                        if ($fullPath -like "$desktopPath*") {
                            $allFindings += [PSCustomObject]@{
                                Source = 'RecycleBin'
                                DeletedTime = $deletedDate
                                FilePath = $fullPath
                                FileName = $itemName
                                Size = $itemSize
                                Status = '✓ RECOVERABLE'
                                Details = 'File is in Recycle Bin - can be restored'
                            }
                            $rbCount++
                        }
                    }
                }
            }
        }
        
        if ($rbCount -gt 0) {
            Write-Host "  ✓ Found $rbCount file(s) - THESE CAN BE RECOVERED!" -ForegroundColor Green
        } else {
            Write-Host "  - No files found in Recycle Bin" -ForegroundColor Gray
        }
        
        if ($shell) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
    } catch {
        Write-Host "  ✗ Error accessing Recycle Bin: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # SOURCE 2: USN JOURNAL (Critical for permanent deletions)
    # ============================================
    Write-Host "`n[2/5] Checking USN Change Journal..." -ForegroundColor Cyan
    $usnCount = 0
    
    try {
        # Query USN journal for the C: drive
        $usnData = fsutil usn readjournal C: csv | ConvertFrom-Csv
        
        $usnDeletions = $usnData | Where-Object {
            try {
                $timestamp = [DateTime]::Parse($_.TimeStamp)
                
                ($_.Name -like "$desktopPath*") -and
                ($_.Reason -match 'FILE_DELETE|RENAME_OLD_NAME|DATA_TRUNCATION') -and
                ($timestamp -ge $Start) -and 
                ($timestamp -le $End)
            } catch {
                $false
            }
        }
        
        foreach ($record in $usnDeletions) {
            $timestamp = [DateTime]::Parse($record.TimeStamp)
            $fileName = Split-Path $record.Name -Leaf
            
            $allFindings += [PSCustomObject]@{
                Source = 'USN_Journal'
                DeletedTime = $timestamp
                FilePath = $record.Name
                FileName = $fileName
                Size = 'Unknown'
                Status = 'Permanently Deleted'
                Details = "Reason: $($record.Reason)"
            }
            $usnCount++
        }
        
        if ($usnCount -gt 0) {
            Write-Host "  ✓ Found $usnCount deletion record(s)" -ForegroundColor Green
        } else {
            Write-Host "  - No deletion records found (may have aged out of journal)" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  ✗ Error reading USN Journal: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # SOURCE 3: SHADOW COPIES (Volume Snapshots)
    # ============================================
    Write-Host "`n[3/5] Checking Shadow Copies..." -ForegroundColor Cyan
    $shadowCount = 0
    
    try {
        $shadowCopies = Get-WmiObject Win32_ShadowCopy | 
            Where-Object { 
                $installDate = $_.InstallDate
                if ($installDate) {
                    $shadowDate = [Management.ManagementDateTimeConverter]::ToDateTime($installDate)
                    ($shadowDate -ge $Start.AddDays(-7)) -and ($shadowDate -le $End)
                } else {
                    $false
                }
            } | Sort-Object InstallDate -Descending
        
        foreach ($shadow in $shadowCopies) {
            $shadowDate = [Management.ManagementDateTimeConverter]::ToDateTime($shadow.InstallDate)
            $shadowPath = "$($shadow.DeviceObject)\Users\$Username\Desktop"
            
            if (Test-Path $shadowPath) {
                $files = Get-ChildItem $shadowPath -Force -ErrorAction SilentlyContinue
                
                $allFindings += [PSCustomObject]@{
                    Source = 'ShadowCopy'
                    DeletedTime = $shadowDate
                    FilePath = $shadowPath
                    FileName = "Shadow Copy available"
                    Size = "$($files.Count) items"
                    Status = '✓ RECOVERABLE from snapshot'
                    Details = "Snapshot ID: $($shadow.ID)"
                }
                $shadowCount++
            }
        }
        
        if ($shadowCount -gt 0) {
            Write-Host "  ✓ Found $shadowCount shadow copy snapshot(s) - files may be recoverable!" -ForegroundColor Green
        } else {
            Write-Host "  - No shadow copies found (feature may not be enabled)" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  ✗ Error checking shadow copies: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # SOURCE 4: SECURITY EVENT LOG (If auditing enabled)
    # ============================================
    Write-Host "`n[4/5] Checking Security Event Log..." -ForegroundColor Cyan
    $secCount = 0
    
    try {
        $secEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            Id = 4663, 4660  # Object access, Object deleted
            StartTime = $Start
            EndTime = $End
        } -ErrorAction SilentlyContinue
        
        foreach ($event in $secEvents) {
            $xml = [xml]$event.ToXml()
            $eventData = @{}
            
            foreach ($data in $xml.Event.EventData.Data) {
                $eventData[$data.Name] = $data.'#text'
            }
            
            $objectName = $eventData['ObjectName']
            
            if ($objectName -and ($objectName -like "$desktopPath*")) {
                $allFindings += [PSCustomObject]@{
                    Source = 'SecurityLog'
                    DeletedTime = $event.TimeCreated
                    FilePath = $objectName
                    FileName = Split-Path $objectName -Leaf
                    Size = 'Unknown'
                    Status = 'Access Logged'
                    Details = "User: $($eventData['SubjectUserName'])"
                }
                $secCount++
            }
        }
        
        if ($secCount -gt 0) {
            Write-Host "  ✓ Found $secCount security event(s)" -ForegroundColor Green
        } else {
            Write-Host "  - No security events found (auditing may not have been enabled)" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  ✗ Error reading Security Log: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # SOURCE 5: POWERSHELL HISTORY
    # ============================================
    Write-Host "`n[5/5] Checking PowerShell History..." -ForegroundColor Cyan
    $psCount = 0
    
    try {
        $historyPath = "C:\Users\$Username\AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt"
        
        if (Test-Path $historyPath) {
            $history = Get-Content $historyPath | Where-Object {
                $_ -match 'Remove-Item|del |rm |rmdir' -and 
                $_ -match 'Desktop'
            }
            
            foreach ($cmd in $history) {
                $allFindings += [PSCustomObject]@{
                    Source = 'PS_History'
                    DeletedTime = 'Unknown'
                    FilePath = 'See Details'
                    FileName = 'PowerShell Command'
                    Size = 'N/A'
                    Status = 'Command Found'
                    Details = $cmd
                }
                $psCount++
            }
            
            if ($psCount -gt 0) {
                Write-Host "  ✓ Found $psCount PowerShell deletion command(s)" -ForegroundColor Green
            } else {
                Write-Host "  - No PowerShell deletion commands found" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  ✗ Error reading PowerShell history: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # ============================================
    # RETURN RESULTS
    # ============================================
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "ANALYSIS COMPLETE" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Yellow
    
    return $allFindings
}

# ============================================
# EXECUTE ANALYSIS
# ============================================

try {
    if ($isLocal) {
        Write-Host "Analyzing local machine..." -ForegroundColor Cyan
        $results = & $scriptBlock -Username $username -Start $start -End $end
    } else {
        Write-Host "Connecting to $computer..." -ForegroundColor Cyan
        $results = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock `
            -ArgumentList $username, $start, $end -ErrorAction Stop
    }
    
    # ============================================
    # DISPLAY RESULTS
    # ============================================
    
    if ($results.Count -eq 0) {
        Write-Host "`n⚠ No deleted files found in any source." -ForegroundColor Yellow
        Write-Host "`nPossible reasons:" -ForegroundColor Gray
        Write-Host "  • Files were deleted outside the specified time range" -ForegroundColor Gray
        Write-Host "  • USN journal has already aged out the records" -ForegroundColor Gray
        Write-Host "  • Files were deleted from a different location" -ForegroundColor Gray
        Write-Host "  • Recycle Bin was emptied" -ForegroundColor Gray
    } else {
        Write-Host "`n✓ FOUND $($results.Count) RECORD(S):`n" -ForegroundColor Green
        
        # Group by recoverability
        $recoverable = $results | Where-Object { $_.Status -like '*RECOVERABLE*' }
        $permanent = $results | Where-Object { $_.Status -notlike '*RECOVERABLE*' }
        
        if ($recoverable) {
            Write-Host "═══ RECOVERABLE FILES ($($recoverable.Count)) ═══" -ForegroundColor Green
            $recoverable | Format-Table -AutoSize
        }
        
        if ($permanent) {
            Write-Host "═══ PERMANENTLY DELETED ($($permanent.Count)) ═══" -ForegroundColor Red
            $permanent | Format-Table -AutoSize
        }
        
        # Export all results
        $results | Sort-Object DeletedTime -Descending | 
            Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outputFile -Force
        
        Write-Host "`n✓ Full report saved to:" -ForegroundColor Green
        Write-Host "  $outputFile`n" -ForegroundColor Cyan
        
        # ============================================
        # RECOVERY RECOMMENDATIONS
        # ============================================
        
        if ($recoverable) {
            Write-Host "═══ RECOVERY OPTIONS ═══" -ForegroundColor Yellow
            Write-Host "`n1. Recycle Bin Recovery:" -ForegroundColor Cyan
            Write-Host "   • Open Recycle Bin on $computer"
            Write-Host "   • Right-click deleted files → 'Restore'"
            
            $shadowCopies = $recoverable | Where-Object { $_.Source -eq 'ShadowCopy' }
            if ($shadowCopies) {
                Write-Host "`n2. Shadow Copy Recovery:" -ForegroundColor Cyan
                Write-Host "   • Right-click Desktop folder → 'Properties' → 'Previous Versions'"
                Write-Host "   • Select a snapshot from the date range"
                Write-Host "   • Click 'Open' to browse deleted files"
                Write-Host "   • Copy files back to Desktop"
            }
        }
        
        if ($permanent -and -not $recoverable) {
            Write-Host "═══ PERMANENT DELETION DETECTED ═══" -ForegroundColor Red
            Write-Host "`n⚠ Files were permanently deleted (bypassed Recycle Bin or emptied)" -ForegroundColor Yellow
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "  1. Check if shadow copies/backups exist (see report above)"
            Write-Host "  2. Use third-party recovery tools (Recuva, PhotoRec, etc.)"
            Write-Host "  3. Contact IT/Security if this is a security incident"
        }
    }
    
} catch {
    Write-Host "`n✗ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}

Write-Host "`nForensic analysis completed.`n" -ForegroundColor Green
