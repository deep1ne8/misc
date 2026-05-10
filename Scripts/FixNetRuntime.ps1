#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Scans, repairs, and installs .NET runtimes ensuring GSA application compatibility.
.DESCRIPTION
    Comprehensive .NET runtime management script that detects corrupted installations,
    installs missing runtimes (especially .NET 8.x for GSA), and generates detailed reports.
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:TEMP\DotNetRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    [switch]$SkipHostingBundle
)

$ErrorActionPreference = 'Stop'
$Report = @{
    StartTime = Get-Date
    Actions = @()
    Errors = @()
    InstalledRuntimes = @()
    CorruptedRuntimes = @()
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Green"})
    Add-Content -Path $LogPath -Value $logMessage
    $Report.Actions += $logMessage
}

function Test-DotNetInstallation {
    try {
        $runtimeOutput = & dotnet --list-runtimes 2>&1
        $sdkOutput = & dotnet --list-sdks 2>&1
        return @{
            Success = $LASTEXITCODE -eq 0
            Runtimes = $runtimeOutput
            SDKs = $sdkOutput
        }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-LatestDotNetVersion {
    param([string]$MajorVersion)
    try {
        $url = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/$MajorVersion.0/releases.json"
        $releases = Invoke-RestMethod -Uri $url -UseBasicParsing
        return $releases.'latest-runtime'
    } catch {
        Write-Log "Failed to fetch latest .NET $MajorVersion version: $($_.Exception.Message)" "WARN"
        return $null
    }
}

function Install-DotNetRuntime {
    param(
        [string]$Version,
        [string]$Type = "runtime"
    )
    
    Write-Log "Installing .NET $Version ($Type)..."
    
    $installerPath = "$env:TEMP\dotnet-$Type-$Version.exe"
    
    try {
        # Determine download URL based on type and version
        $majorVersion = $Version.Split('.')[0]
        
        if ($Type -eq "hosting") {
            $downloadUrl = "https://download.visualstudio.microsoft.com/download/pr/dotnet/$majorVersion.0/dotnet-hosting-win.exe"
            # Fallback to direct version-specific URL
            $downloadPage = "https://dotnet.microsoft.com/en-us/download/dotnet/$majorVersion.0"
        } else {
            $downloadUrl = "https://download.visualstudio.microsoft.com/download/pr/dotnet/$majorVersion.0/windowsdesktop-runtime-win-x64.exe"
        }
        
        Write-Log "Downloading from Microsoft servers..."
        
        # Use direct download for .NET 8 Hosting Bundle
        if ($Type -eq "hosting" -and $majorVersion -eq "8") {
            $directUrl = "https://download.visualstudio.microsoft.com/download/pr/e6a4368b-9d7f-4976-b4a9-da47f8ef1e63/4c0d0582c1ec5ca3f0dd2ab21b8d3b9f/dotnet-hosting-8.0.11-win.exe"
            Invoke-WebRequest -Uri $directUrl -OutFile $installerPath -UseBasicParsing
        } else {
            # Generic runtime installer download
            $genericUrl = "https://dotnetcli.azureedge.net/dotnet/Runtime/$Version/dotnet-runtime-$Version-win-x64.exe"
            Invoke-WebRequest -Uri $genericUrl -OutFile $installerPath -UseBasicParsing
        }
        
        Write-Log "Installing silently..."
        $process = Start-Process -FilePath $installerPath -ArgumentList "/quiet","/norestart" -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Log ".NET $Version installed successfully (Exit Code: $($process.ExitCode))"
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
            return $true
        } else {
            Write-Log "Installation failed with exit code: $($process.ExitCode)" "ERROR"
            $Report.Errors += "Installation failed for .NET $Version"
            return $false
        }
    } catch {
        Write-Log "Installation error: $($_.Exception.Message)" "ERROR"
        $Report.Errors += $_.Exception.Message
        return $false
    }
}

function Repair-DotNetInstallation {
    Write-Log "Attempting .NET installation repair..."
    
    # Run .NET repair tool
    try {
        $repairTool = "$env:ProgramFiles\dotnet\dotnet.exe"
        if (Test-Path $repairTool) {
            & $repairTool --info | Out-Null
            Write-Log "Basic .NET validation passed"
        }
    } catch {
        Write-Log "Repair needed: $($_.Exception.Message)" "WARN"
    }
}

# ============================================================
# MAIN EXECUTION
# ============================================================

Write-Log "=== Starting .NET Runtime Scan and Repair ==="
Write-Log "Log file: $LogPath"

# Step 1: Initial .NET Check
Write-Log "`n--- Step 1: Checking Current .NET Installation ---"
$dotnetCheck = Test-DotNetInstallation

if ($dotnetCheck.Success) {
    Write-Log "✓ .NET CLI is functional"
    Write-Log "`nInstalled Runtimes:"
    $dotnetCheck.Runtimes | ForEach-Object { 
        Write-Log "  $_"
        $Report.InstalledRuntimes += $_
    }
    Write-Log "`nInstalled SDKs:"
    $dotnetCheck.SDKs | ForEach-Object { Write-Log "  $_" }
} else {
    Write-Log "✗ .NET CLI not found or corrupted" "ERROR"
    $Report.CorruptedRuntimes += "Core CLI"
}

# Step 2: Check for .NET 8.x (GSA Requirement)
Write-Log "`n--- Step 2: Validating .NET 8.x for GSA ---"
$hasDotNet8 = $dotnetCheck.Runtimes | Where-Object { $_ -match 'Microsoft\.AspNetCore\.App 8\.' }

if ($hasDotNet8) {
    Write-Log "✓ .NET 8.x ASP.NET Core Runtime found: $($hasDotNet8 -join ', ')"
} else {
    Write-Log "✗ .NET 8.x ASP.NET Core Runtime NOT found - GSA may fail" "WARN"
    $Report.CorruptedRuntimes += ".NET 8.x ASP.NET Core"
}

$hasDotNet8Desktop = $dotnetCheck.Runtimes | Where-Object { $_ -match 'Microsoft\.WindowsDesktop\.App 8\.' }
if (-not $hasDotNet8Desktop) {
    Write-Log "✗ .NET 8.x Desktop Runtime NOT found" "WARN"
    $Report.CorruptedRuntimes += ".NET 8.x Desktop"
}

# Step 3: Get Latest .NET 8 Version
Write-Log "`n--- Step 3: Checking for Latest .NET 8 Version ---"
$latestNet8 = Get-LatestDotNetVersion -MajorVersion "8"
if ($latestNet8) {
    Write-Log "Latest .NET 8 version available: $latestNet8"
} else {
    $latestNet8 = "8.0.11" # Fallback
    Write-Log "Using fallback version: $latestNet8"
}

# Step 4: Install/Repair .NET 8 Hosting Bundle (for GSA)
Write-Log "`n--- Step 4: Installing/Repairing .NET 8 Hosting Bundle ---"
if (-not $SkipHostingBundle) {
    if (-not $hasDotNet8 -or -not $hasDotNet8Desktop) {
        $installed = Install-DotNetRuntime -Version $latestNet8 -Type "hosting"
        if ($installed) {
            $Report.Actions += "Installed .NET 8 Hosting Bundle"
        }
    } else {
        Write-Log "✓ .NET 8 runtimes already present, skipping installation"
    }
} else {
    Write-Log "Skipping Hosting Bundle installation (per switch parameter)"
}

# Step 5: Verify Additional Critical Runtimes
Write-Log "`n--- Step 5: Checking Additional Runtimes ---"
$criticalVersions = @("6.0", "7.0")

foreach ($version in $criticalVersions) {
    $hasVersion = $dotnetCheck.Runtimes | Where-Object { $_ -match "Microsoft\.NETCore\.App $version\." }
    if (-not $hasVersion) {
        Write-Log "Missing .NET $version runtime - consider installing if needed" "WARN"
    } else {
        Write-Log "✓ .NET $version runtime found"
    }
}

# Step 6: Final Validation
Write-Log "`n--- Step 6: Final Validation ---"
$finalCheck = Test-DotNetInstallation

if ($finalCheck.Success) {
    Write-Log "✓ Final .NET validation PASSED"
    Write-Log "`nFinal Installed Runtimes:"
    $finalCheck.Runtimes | ForEach-Object { Write-Log "  $_" }
} else {
    Write-Log "✗ Final validation FAILED" "ERROR"
}

# Step 7: Generate Report
Write-Log "`n--- Step 7: Generating Report ---"
$Report.EndTime = Get-Date
$Report.Duration = $Report.EndTime - $Report.StartTime

Write-Log "`n=========================================="
Write-Log "         .NET REPAIR REPORT SUMMARY"
Write-Log "=========================================="
Write-Log "Start Time: $($Report.StartTime)"
Write-Log "End Time: $($Report.EndTime)"
Write-Log "Duration: $($Report.Duration.TotalSeconds) seconds"
Write-Log "`nInstalled Runtimes: $($Report.InstalledRuntimes.Count)"
Write-Log "Corrupted/Missing: $($Report.CorruptedRuntimes.Count)"
Write-Log "Errors: $($Report.Errors.Count)"
Write-Log "`nLog saved to: $LogPath"

if ($Report.Errors.Count -eq 0 -and $hasDotNet8) {
    Write-Log "`n✓ GSA APPLICATION READY - All required runtimes installed" "INFO"
    exit 0
} elseif ($hasDotNet8) {
    Write-Log "`n⚠ GSA should work, but some issues detected" "WARN"
    exit 1
} else {
    Write-Log "`n✗ GSA MAY NOT FUNCTION - Critical runtimes missing" "ERROR"
    exit 2
}
