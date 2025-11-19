#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads and installs Windows 11 24H2 cumulative updates
.DESCRIPTION
    Downloads KB5043080 and KB5064081 updates and installs them
#>

[CmdletBinding()]
param(
    [string]$DownloadPath = "$env:TEMP\WindowsUpdates"
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Green"})
}

function Get-FileHash256 {
    param([string]$FilePath)
    return (Get-FileHash -Path $FilePath -Algorithm SHA256).Hash
}

function Download-WithProgress {
    param([string]$Url, [string]$OutputPath)
    
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
        
        Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -SourceIdentifier WebClient.DownloadProgressChanged -Action {
            $percent = $EventArgs.ProgressPercentage
            $received = [math]::Round($EventArgs.BytesReceived / 1MB, 2)
            $total = [math]::Round($EventArgs.TotalBytesToReceive / 1MB, 2)
            Write-Progress -Activity "Downloading" -Status "$received MB / $total MB" -PercentComplete $percent
        } | Out-Null
        
        Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -SourceIdentifier WebClient.DownloadFileCompleted | Out-Null
        
        $webClient.DownloadFileAsync($Url, $OutputPath)
        
        while ($webClient.IsBusy) {
            Start-Sleep -Milliseconds 100
        }
        
        Write-Progress -Activity "Downloading" -Completed
        
    } finally {
        Unregister-Event -SourceIdentifier WebClient.DownloadProgressChanged -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier WebClient.DownloadFileCompleted -ErrorAction SilentlyContinue
        if ($webClient) { $webClient.Dispose() }
    }
}

function Install-MSU {
    param([string]$MsuPath, [string]$KB)
    
    Write-Log "Installing $KB..."
    $process = Start-Process -FilePath "wusa.exe" -ArgumentList "`"$MsuPath`" /quiet /norestart" -Wait -PassThru -NoNewWindow
    
    switch ($process.ExitCode) {
        0 { Write-Log "$KB installed successfully"; return $true }
        3010 { Write-Log "$KB installed (reboot required)" "WARN"; return $true }
        2359302 { Write-Log "$KB already installed" "WARN"; return $true }
        default { Write-Log "$KB installation failed (Exit code: $($process.ExitCode))" "ERROR"; return $false }
    }
}

# Update definitions
$updates = @(
    @{
        KB = "KB5043080"
        FileName = "windows11.0-kb5043080-x64_953449672073f8fb99badb4cc6d5d7849b9c83e8.msu"
        SHA256 = "8196328210101AF11C7290F87B13FF05BF3489B22208263EA03BCC1D1F26A640"
        Url = "https://catalog.s.download.windowsupdate.com/c/msdownload/update/software/updt/2024/08/windows11.0-kb5043080-x64_953449672073f8fb99badb4cc6d5d7849b9c83e8.msu"
    },
    @{
        KB = "KB5064081"
        FileName = "windows11.0-kb5064081-x64_a1096145ded3adfc26f8f23442281533429f0e38.msu"
        SHA256 = "C4BD7AFC0783ECDDD5438866909483F3DE0D0B3F18FF00A677D157E44FF50401"
        Url = "https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/updt/2025/08/windows11.0-kb5064081-x64_a1096145ded3adfc26f8f23442281533429f0e38.msu"
    }
)

try {
    Write-Log "Windows 11 24H2 Cumulative Update Installer"
    Write-Log "============================================"
    
    # Verify OS version
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $buildNumber = $osInfo.BuildNumber
    Write-Log "Current OS: $($osInfo.Caption) Build $buildNumber"
    
    if ($buildNumber -ne "26100") {
        Write-Log "This update is for Windows 11 24H2 (Build 26100). Current build: $buildNumber" "WARN"
        $continue = Read-Host "Continue anyway? (Y/N)"
        if ($continue -ne 'Y') { exit 0 }
    }
    
    # Create download directory
    if (-not (Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
        Write-Log "Created download directory: $DownloadPath"
    }
    
    $rebootRequired = $false
    
    foreach ($update in $updates) {
        Write-Log ""
        Write-Log "Processing $($update.KB)..."
        
        $filePath = Join-Path $DownloadPath $update.FileName
        
        # Check if already downloaded and verified
        if (Test-Path $filePath) {
            Write-Log "File exists, verifying hash..."
            $hash = Get-FileHash256 -FilePath $filePath
            
            if ($hash -eq $update.SHA256) {
                Write-Log "Hash verified successfully"
            } else {
                Write-Log "Hash mismatch, re-downloading..." "WARN"
                Remove-Item $filePath -Force
            }
        }
        
        # Download if needed
        if (-not (Test-Path $filePath)) {
            Write-Log "Downloading $($update.KB) from Microsoft Update Catalog..."
            Download-WithProgress -Url $update.Url -OutputPath $filePath
            
            Write-Log "Verifying download integrity..."
            $hash = Get-FileHash256 -FilePath $filePath
            
            if ($hash -ne $update.SHA256) {
                throw "Hash verification failed for $($update.KB). Expected: $($update.SHA256), Got: $hash"
            }
            Write-Log "Download verified successfully"
        }
        
        # Install update
        $installed = Install-MSU -MsuPath $filePath -KB $update.KB
        if ($installed) { $rebootRequired = $true }
    }
    
    Write-Log ""
    Write-Log "============================================"
    Write-Log "All updates processed successfully"
    
    if ($rebootRequired) {
        Write-Log "SYSTEM REBOOT REQUIRED" "WARN"
        $reboot = Read-Host "Reboot now? (Y/N)"
        if ($reboot -eq 'Y') {
            Write-Log "Initiating restart in 30 seconds..."
            shutdown /r /t 30 /c "Restarting to complete Windows updates"
        } else {
            Write-Log "Please restart your computer to complete the installation"
        }
    }
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
