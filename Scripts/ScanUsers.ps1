Get-ChildItem "C:\Users" -Directory | ForEach-Object {
    $userPath = "$($_.FullName)\NTUSER.DAT"
    if (Test-Path $userPath) {
        $sid = $_.Name
        reg load "HKU\TempHive_$sid" $userPath | Out-Null
        $key = "Registry::HKEY_USERS\TempHive_$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        if (Test-Path $key) {
            $props = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                UserName  = $_.Name
                Desktop   = $props.Desktop
                Documents = $props.Personal
            }
        }
        reg unload "HKU\TempHive_$sid" | Out-Null
    }
} | Where-Object { $_.Desktop -or $_.Documents } | Format-Table -AutoSize
