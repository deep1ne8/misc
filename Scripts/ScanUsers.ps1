Get-ChildItem 'HKU:' | ForEach-Object { 
    $sid = $_.PSChildName
    $key = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    if (Test-Path $key) {
        [PSCustomObject]@{
            SID        = $sid
            Desktop    = (Get-ItemProperty -Path $key -Name 'Desktop' -ErrorAction SilentlyContinue).Desktop
            Documents  = (Get-ItemProperty -Path $key -Name 'Personal' -ErrorAction SilentlyContinue).Personal
        }
    }
} | Where-Object { $_.Desktop -or $_.Documents } | Format-Table -AutoSize
