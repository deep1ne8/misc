<#
.SYNOPSIS
    Mitigates PrintNightmare while maintaining Seagull driver functionality
.DESCRIPTION
    Balances security (PrintNightmare protection) with functionality (Seagull driver installation)
#>

[CmdletBinding()]
param(
    [switch]$SecureMode,  # Enforces strict PrintNightmare protection (may break Seagull)
    [switch]$CompatibilityMode  # Allows Seagull drivers to work (less secure)
)

$ErrorActionPreference = "Stop"
$LogPath = "C:\Logs\PrintNightmare_Seagull_Fix.log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogPath -Value $LogMessage -Force
}

# Create log directory
New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue | Out-Null

Write-Log "=== PrintNightmare Mitigation with Seagull Driver Support ===" "INFO"
Write-Log "Mode: $(if($SecureMode){'SECURE'}elseif($CompatibilityMode){'COMPATIBILITY'}else{'BALANCED'})"

# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Log "ERROR: Script must run as Administrator" "ERROR"
    throw "Run as Administrator"
}

# Registry paths
$PrintersPoliciesPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers"
$PointAndPrintPath = "$PrintersPoliciesPath\PointAndPrint"

# Create registry paths if they don't exist
if (!(Test-Path $PrintersPoliciesPath)) {
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT" -Name "Printers" -Force | Out-Null
    Write-Log "Created Printers policy path"
}

if (!(Test-Path $PointAndPrintPath)) {
    New-Item -Path $PrintersPoliciesPath -Name "PointAndPrint" -Force | Out-Null
    Write-Log "Created PointAndPrint policy path"
}

# ===== PRINTNIGHTMARE MITIGATION =====
Write-Log "Applying PrintNightmare mitigation settings..."

try {
    # Critical: Disable remote RPC endpoint (blocks PrintNightmare RCE)
    Set-ItemProperty -Path $PrintersPoliciesPath -Name "RegisterSpoolerRemoteRpcEndPoint" -Value 2 -Type DWord
    Write-Log "Set RegisterSpoolerRemoteRpcEndPoint = 2 (RPC Endpoint disabled)"
    
    # Only load printer drivers as Admin (blocks LPE)
    Set-ItemProperty -Path $PrintersPoliciesPath -Name "RestrictDriverInstallationToAdministrators" -Value 1 -Type DWord
    Write-Log "Set RestrictDriverInstallationToAdministrators = 1"
    
} catch {
    Write-Log "ERROR: Failed to apply PrintNightmare settings - $_" "ERROR"
}

# ===== SEAGULL DRIVER COMPATIBILITY =====
if ($CompatibilityMode) {
    Write-Log "Applying COMPATIBILITY mode for Seagull drivers..." "WARN"
    Write-Log "WARNING: This reduces security but enables Seagull driver installation" "WARN"
    
    try {
        # Allow non-admin driver installation (LESS SECURE)
        Set-ItemProperty -Path $PointAndPrintPath -Name "RestrictDriverInstallationToAdministrators" -Value 0 -Type DWord
        Write-Log "Set RestrictDriverInstallationToAdministrators = 0 (COMPATIBILITY)"
        
        # Disable elevation prompts
        Set-ItemProperty -Path $PointAndPrintPath -Name "NoWarningNoElevationOnInstall" -Value 1 -Type DWord
        Write-Log "Set NoWarningNoElevationOnInstall = 1"
        
        # Disable update prompts
        Set-ItemProperty -Path $PointAndPrintPath -Name "UpdatePromptSettings" -Value 0 -Type DWord
        Write-Log "Set UpdatePromptSettings = 0"
        
        # Allow package-aware drivers
        Set-ItemProperty -Path $PointAndPrintPath -Name "PackagePointAndPrintOnly" -Value 0 -Type DWord
        Write-Log "Set PackagePointAndPrintOnly = 0"
        
        # Trust any server
        Set-ItemProperty -Path $PointAndPrintPath -Name "PackagePointAndPrintServerList" -Value 1 -Type DWord
        Write-Log "Set PackagePointAndPrintServerList = 1"
        
    } catch {
        Write-Log "ERROR: Failed to apply compatibility settings - $_" "ERROR"
    }
    
} elseif ($SecureMode) {
    Write-Log "Applying SECURE mode - Seagull drivers may not work..." "WARN"
    
    try {
        # Enforce strict security (may break Seagull)
        Set-ItemProperty -Path $PointAndPrintPath -Name "RestrictDriverInstallationToAdministrators" -Value 1 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "NoWarningNoElevationOnInstall" -Value 0 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "UpdatePromptSettings" -Value 1 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "PackagePointAndPrintOnly" -Value 1 -Type DWord
        Write-Log "Applied strict security settings"
        
    } catch {
        Write-Log "ERROR: Failed to apply secure settings - $_" "ERROR"
    }
    
} else {
    Write-Log "Applying BALANCED mode (recommended)..."
    
    try {
        # Balanced: Require elevation but allow trusted drivers
        Set-ItemProperty -Path $PointAndPrintPath -Name "RestrictDriverInstallationToAdministrators" -Value 0 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "NoWarningNoElevationOnInstall" -Value 0 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "UpdatePromptSettings" -Value 2 -Type DWord  # Prompt for updates
        Set-ItemProperty -Path $PointAndPrintPath -Name "PackagePointAndPrintOnly" -Value 0 -Type DWord
        Set-ItemProperty -Path $PointAndPrintPath -Name "TrustedServers" -Value 0 -Type DWord
        Write-Log "Applied balanced security settings"
        
    } catch {
        Write-Log "ERROR: Failed to apply balanced settings - $_" "ERROR"
    }
}

# ===== SEAGULL-SPECIFIC FIXES =====
Write-Log "Applying Seagull driver specific fixes..."

# Configure Windows Update to allow signed content from intranet
try {
    $WUPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    if (!(Test-Path $WUPath)) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "WindowsUpdate" -Force | Out-Null
    }
    Set-ItemProperty -Path $WUPath -Name "AcceptTrustedPublisherCerts" -Value 1 -Type DWord
    Write-Log "Enabled AcceptTrustedPublisherCerts for Seagull certificates"
} catch {
    Write-Log "WARNING: Could not configure Windows Update settings - $_" "WARN"
}

# ===== FIREWALL RULES FOR SEAGULL TCP =====
Write-Log "Configuring firewall for Seagull TCP ports..."

try {
    # Port 5130 (Seagull 2017.1+)
    $Rule5130 = Get-NetFirewallRule -DisplayName "Seagull TCP 5130" -ErrorAction SilentlyContinue
    if (-not $Rule5130) {
        New-NetFirewallRule -DisplayName "Seagull TCP 5130" `
            -Direction Inbound -Protocol TCP -LocalPort 5130 `
            -Action Allow -Profile Domain,Private -Enabled True | Out-Null
        Write-Log "Created firewall rule for TCP 5130"
    } else {
        Write-Log "Firewall rule for TCP 5130 already exists"
    }
    
    # Port 6160 (Legacy Seagull)
    $Rule6160 = Get-NetFirewallRule -DisplayName "Seagull TCP 6160" -ErrorAction SilentlyContinue
    if (-not $Rule6160) {
        New-NetFirewallRule -DisplayName "Seagull TCP 6160" `
            -Direction Inbound -Protocol TCP -LocalPort 6160 `
            -Action Allow -Profile Domain,Private -Enabled True | Out-Null
        Write-Log "Created firewall rule for TCP 6160"
    } else {
        Write-Log "Firewall rule for TCP 6160 already exists"
    }
    
} catch {
    Write-Log "WARNING: Firewall rules may need manual configuration - $_" "WARN"
}

# ===== PRINT SPOOLER SERVICE =====
Write-Log "Checking Print Spooler service status..."

try {
    $SpoolerService = Get-Service -Name Spooler
    if ($SpoolerService.Status -ne "Running") {
        Write-Log "WARNING: Print Spooler is not running. Starting..." "WARN"
        Start-Service Spooler
        Write-Log "Print Spooler started"
    } else {
        Write-Log "Print Spooler is running - will restart to apply changes"
        Restart-Service Spooler -Force
        Start-Sleep -Seconds 3
        Write-Log "Print Spooler restarted"
    }
} catch {
    Write-Log "ERROR: Failed to manage Print Spooler - $_" "ERROR"
}

# ===== VERIFICATION =====
Write-Log "=== Configuration Verification ==="

try {
    $RegValues = @(
        @{Path=$PrintersPoliciesPath; Name="RegisterSpoolerRemoteRpcEndPoint"; Expected=2},
        @{Path=$PointAndPrintPath; Name="RestrictDriverInstallationToAdministrators"},
        @{Path=$PointAndPrintPath; Name="NoWarningNoElevationOnInstall"},
        @{Path=$PointAndPrintPath; Name="UpdatePromptSettings"}
    )
    
    foreach ($RegValue in $RegValues) {
        $Value = Get-ItemProperty -Path $RegValue.Path -Name $RegValue.Name -ErrorAction SilentlyContinue
        if ($Value) {
            Write-Log "$($RegValue.Name) = $($Value.$($RegValue.Name))"
        }
    }
} catch {
    Write-Log "WARNING: Could not verify all settings - $_" "WARN"
}

Write-Log "=== Script Completed ===" "INFO"
Write-Log "Log saved to: $LogPath"

# Display summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CONFIGURATION COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Mode Applied: $(if($SecureMode){'SECURE (High Security)'}elseif($CompatibilityMode){'COMPATIBILITY (Seagull-Friendly)'}else{'BALANCED (Recommended)'})" -ForegroundColor Yellow
Write-Host "`nNext Steps:" -ForegroundColor White
Write-Host "1. Test Seagull driver installation" -ForegroundColor White
Write-Host "2. If 0x00000bcb persists, install Seagull certificate via GPO" -ForegroundColor White
Write-Host "3. Update to Seagull driver version 2022+ " -ForegroundColor White
Write-Host "4. Review log: $LogPath" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan
