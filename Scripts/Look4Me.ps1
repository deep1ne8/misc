# Robust Deleted File Search Script
# Run elevated for full access

# Prompt user
$choice = Read-Host "Select time range (Day, Week, Month, 6months, Year)"
$filename = Read-Host "Enter part or full filename to search for (no wildcards needed)"

# Convert time range into cutoff date
switch -Regex ($choice.ToLower()) {
    "day"      { $cutoff = (Get-Date).AddDays(-1) }
    "week"     { $cutoff = (Get-Date).AddDays(-7) }
    "month"    { $cutoff = (Get-Date).AddMonths(-1) }
    "6months"  { $cutoff = (Get-Date).AddMonths(-6) }
    "year"     { $cutoff = (Get-Date).AddYears(-1) }
    default    { Write-Host "Invalid option, defaulting to 1 month."; $cutoff = (Get-Date).AddMonths(-1) }
}

Write-Host "`n[+] Searching for deleted files matching '$filename' since $cutoff...`n" -ForegroundColor Cyan

# Prepare results array
$results = @()

# 1. Search Recycle Bin
$recycle = "$env:SystemDrive`$\`$Recycle.Bin"
if (Test-Path $recycle) {
    Get-ChildItem $recycle -Recurse -ErrorAction SilentlyContinue | 
        Where-Object { $_.Name -match $filename -and $_.LastWriteTime -gt $cutoff } |
        ForEach-Object {
            $results += [PSCustomObject]@{
                Source  = "RecycleBin"
                File    = $_.Name
                Path    = $_.FullName
                Deleted = $_.LastWriteTime
            }
        }
}

# 2. Search FileSystem change logs (Event ID 4663 = file deleted)
$logEvents = Get-WinEvent -LogName Security -ErrorAction SilentlyContinue |
    Where-Object { $_.Id -eq 4663 -and $_.TimeCreated -gt $cutoff -and $_.Message -match "Delete" -and $_.Message -match $filename }

foreach ($evt in $logEvents) {
    $results += [PSCustomObject]@{
        Source  = "SecurityLog"
        File    = ($evt.Message -split "`n" | Select-String -Pattern $filename | Select-Object -First 1).ToString()
        Path    = "EventID: $($evt.Id)"
        Deleted = $evt.TimeCreated
    }
}

# 3. Check shadow copies (if enabled)
$vss = Get-WmiObject Win32_ShadowCopy -ErrorAction SilentlyContinue | Where-Object { $_.InstallDate -gt $cutoff }
foreach ($shadow in $vss) {
    $results += [PSCustomObject]@{
        Source  = "ShadowCopy"
        File    = "Snapshot ID: $($shadow.ID)"
        Path    = $shadow.DeviceObject
        Deleted = ([WMI]$shadow.__PATH).InstallDate
    }
}

# Output results
if ($results) {
    $results | Sort-Object Deleted -Descending | Format-Table -AutoSize
    $export = "$env:USERPROFILE\Desktop\DeletedFiles_Report_$(Get-Date -f 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $export -NoTypeInformation
    Write-Host "`nReport exported to: $export" -ForegroundColor Green
} else {
    Write-Host "`nNo deleted files found for '$filename' in the selected period." -ForegroundColor Yellow
}
