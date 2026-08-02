#Requires -Version 5.1

$ErrorActionPreference = 'Continue'

$LogPath = Join-Path $env:TEMP "Liongard-Agent-Removal.log"

function Write-Log {
param(
[string]$Message,
[string]$Level = "INFO"
)

```
$Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message

Write-Host $Entry

try {
    Add-Content -Path $LogPath -Value $Entry -Encoding UTF8 -ErrorAction SilentlyContinue
}
catch {
}
```

}

function Get-LiongardUninstallEntries {
$RegistryPaths = @(
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall*",
"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall*"
)

```
$Entries = foreach ($RegistryPath in $RegistryPaths) {
    Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -match '(?i)Liongard' -or
            $_.Publisher -match '(?i)Liongard' -or
            $_.InstallLocation -match '(?i)Liongard'
        }
}

$Entries |
    Sort-Object PSPath -Unique
```

}

function Stop-LiongardComponents {
Write-Log "Stopping Liongard services and processes."

```
Get-Service -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '(?i)Liongard' -or
        $_.DisplayName -match '(?i)Liongard'
    } |
    ForEach-Object {
        try {
            if ($_.Status -ne 'Stopped') {
                Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
                Write-Log "Stopped service: $($_.Name)"
            }
        }
        catch {
            Write-Log "Unable to stop service $($_.Name): $($_.Exception.Message)" "WARN"
        }
    }

Get-Process -ErrorAction SilentlyContinue |
    Where-Object {
        $_.ProcessName -match '(?i)Liongard'
    } |
    ForEach-Object {
        try {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            Write-Log "Stopped process: $($_.ProcessName), PID $($_.Id)"
        }
        catch {
            Write-Log "Unable to stop process $($_.ProcessName): $($_.Exception.Message)" "WARN"
        }
    }
```

}

function Invoke-LiongardUninstall {
param(
[Parameter(Mandatory)]
[object]$Entry
)

```
$DisplayName = $Entry.DisplayName
$ProductCode = $null

if ($Entry.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
    $ProductCode = $Entry.PSChildName
}

if ($ProductCode) {
    Write-Log "Attempting MSI uninstall for: $DisplayName"
    Write-Log "Product code: $ProductCode"

    try {
        $Process = Start-Process `
            -FilePath "msiexec.exe" `
            -ArgumentList "/x $ProductCode /qn /norestart REBOOT=ReallySuppress" `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        Write-Log "MSI uninstall exit code: $($Process.ExitCode)"

        return $Process.ExitCode
    }
    catch {
        Write-Log "MSI uninstall failed: $($_.Exception.Message)" "ERROR"

        return 1603
    }
}

if ($Entry.QuietUninstallString) {
    $UninstallCommand = $Entry.QuietUninstallString
    Write-Log "Using QuietUninstallString for: $DisplayName"

    try {
        $Process = Start-Process `
            -FilePath "cmd.exe" `
            -ArgumentList "/c $UninstallCommand" `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        Write-Log "Silent uninstall exit code: $($Process.ExitCode)"

        return $Process.ExitCode
    }
    catch {
        Write-Log "Silent uninstall failed: $($_.Exception.Message)" "ERROR"

        return 1603
    }
}

if ($Entry.UninstallString) {
    $UninstallCommand = $Entry.UninstallString

    if ($UninstallCommand -match '(?i)msiexec(\.exe)?') {
        $UninstallCommand = $UninstallCommand `
            -replace '(?i)/I(?=\s|\{)', '/X'

        if ($UninstallCommand -notmatch '(?i)/q') {
            $UninstallCommand += ' /qn'
        }

        if ($UninstallCommand -notmatch '(?i)/norestart') {
            $UninstallCommand += ' /norestart REBOOT=ReallySuppress'
        }
    }

    Write-Log "Using uninstall command for: $DisplayName"

    try {
        $Process = Start-Process `
            -FilePath "cmd.exe" `
            -ArgumentList "/c $UninstallCommand" `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        Write-Log "Uninstall exit code: $($Process.ExitCode)"

        return $Process.ExitCode
    }
    catch {
        Write-Log "Uninstall command failed: $($_.Exception.Message)" "ERROR"

        return 1603
    }
}

Write-Log "No uninstall command found for: $DisplayName" "WARN"

return 1612
```

}

function Remove-LiongardRemnants {
Write-Log "Starting Liongard cleanup."

```
Stop-LiongardComponents

$Services = Get-Service -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '(?i)Liongard' -or
        $_.DisplayName -match '(?i)Liongard'
    }

foreach ($Service in $Services) {
    try {
        & sc.exe delete $Service.Name | Out-Null
        Write-Log "Deleted service: $($Service.Name)"
    }
    catch {
        Write-Log "Unable to delete service $($Service.Name): $($_.Exception.Message)" "WARN"
    }
}

$Folders = @(
    "$env:ProgramFiles\Liongard",
    "$env:ProgramFiles\Liongard Agent",
    "${env:ProgramFiles(x86)}\Liongard",
    "${env:ProgramFiles(x86)}\Liongard Agent",
    "$env:ProgramData\Liongard",
    "$env:ProgramData\LionGard",
    "$env:LOCALAPPDATA\Liongard",
    "$env:APPDATA\Liongard"
) |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_)
    } |
    Select-Object -Unique

foreach ($Folder in $Folders) {
    if (Test-Path $Folder) {
        try {
            Remove-Item -Path $Folder -Recurse -Force -ErrorAction Stop
            Write-Log "Removed folder: $Folder"
        }
        catch {
            Write-Log "Unable to remove folder $Folder: $($_.Exception.Message)" "WARN"
        }
    }
}

$RegistryRoots = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($RegistryRoot in $RegistryRoots) {
    Get-ChildItem -Path $RegistryRoot -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                $Property = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue

                if (
                    $Property.DisplayName -match '(?i)Liongard' -or
                    $Property.Publisher -match '(?i)Liongard' -or
                    $Property.InstallLocation -match '(?i)Liongard'
                ) {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed uninstall registry entry: $($_.PSChildName)"
                }
            }
            catch {
            }
        }
}

$AdditionalRegistryPaths = @(
    "HKLM:\SOFTWARE\Liongard",
    "HKLM:\SOFTWARE\WOW6432Node\Liongard"
)

foreach ($RegistryPath in $AdditionalRegistryPaths) {
    if (Test-Path $RegistryPath) {
        try {
            Remove-Item -Path $RegistryPath -Recurse -Force -ErrorAction Stop
            Write-Log "Removed registry path: $RegistryPath"
        }
        catch {
            Write-Log "Unable to remove registry path $RegistryPath: $($_.Exception.Message)" "WARN"
        }
    }
}
```

}

Write-Log "Starting dynamic Liongard Agent removal."
Write-Log "Log file: $LogPath"

$IsAdministrator = (
[Security.Principal.WindowsPrincipal]
[Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole(
[Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $IsAdministrator) {
Write-Log "Administrator privileges are required." "ERROR"
exit 1
}

$LiongardEntries = @(Get-LiongardUninstallEntries)

if ($LiongardEntries.Count -eq 0) {
Write-Log "No Liongard uninstall entries were found."
Write-Log "Checking for orphaned services, processes, files, and registry entries."

```
Remove-LiongardRemnants

Write-Log "Liongard cleanup completed."
exit 0
```

}

Write-Log "Found $($LiongardEntries.Count) Liongard installation entry or entries."

$UninstallFailed = $false

foreach ($Entry in $LiongardEntries) {
Write-Log "Detected: $($Entry.DisplayName)"

```
Stop-LiongardComponents

$ExitCode = Invoke-LiongardUninstall -Entry $Entry

switch ($ExitCode) {
    0 {
        Write-Log "Liongard uninstall completed successfully."
    }

    3010 {
        Write-Log "Liongard uninstall completed. A restart is required." "WARN"
    }

    1641 {
        Write-Log "Liongard uninstall completed and initiated a restart." "WARN"
    }

    1605 {
        Write-Log "Product is already removed. Performing cleanup." "WARN"
        Remove-LiongardRemnants
    }

    1612 {
        Write-Log "Installer source is unavailable. Performing forced cleanup." "WARN"
        Remove-LiongardRemnants
    }

    default {
        Write-Log "Uninstall returned exit code $ExitCode. Performing cleanup." "WARN"
        $UninstallFailed = $true
        Remove-LiongardRemnants
    }
}
```

}

$RemainingEntries = @(Get-LiongardUninstallEntries)

if ($RemainingEntries.Count -gt 0) {
Write-Log "Liongard uninstall entries remain. Running final cleanup." "WARN"
Remove-LiongardRemnants
}

$RemainingServices = @(
Get-Service -ErrorAction SilentlyContinue |
Where-Object {
$*.Name -match '(?i)Liongard' -or
$*.DisplayName -match '(?i)Liongard'
}
)

$RemainingFolders = @(
"$env:ProgramFiles\Liongard",
"$env:ProgramFiles\Liongard Agent",
"${env:ProgramFiles(x86)}\Liongard",
"${env:ProgramFiles(x86)}\Liongard Agent",
"$env:ProgramData\Liongard",
"$env:ProgramData\LionGard"
) |
Where-Object {
-not [string]::IsNullOrWhiteSpace($*) -and
(Test-Path $*)
}

if (
$RemainingEntries.Count -eq 0 -and
$RemainingServices.Count -eq 0 -and
$RemainingFolders.Count -eq 0
) {
Write-Log "Liongard Agent was removed successfully."
exit 0
}

Write-Log "Liongard cleanup completed, but some remnants may remain." "ERROR"

if ($UninstallFailed) {
exit 1
}

exit 0
