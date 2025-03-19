# Windows 11 23H2 Profile Fix Script
# Run with administrator privileges for complete functionality
# This script attempts to fix common profile loading issues with Windows 11 23H2

# Function to write colored output
function Write-ColorOutput {
    param([string]$message, [string]$color = "White")
    Write-Host $message -ForegroundColor $color
}

Write-ColorOutput "Windows 11 23H2 Profile Fix Tool" "Cyan"
Write-ColorOutput "===========================" "Cyan"
Write-ColorOutput "This script will attempt to fix profile sign-in issues on Windows 11 23H2" "Yellow"
Write-ColorOutput "Please ensure you run this as Administrator" "Yellow"
Write-ColorOutput ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-ColorOutput "ERROR: This script must be run as Administrator!" "Red"
    Write-ColorOutput "Please restart PowerShell as Administrator and try again." "Red"
    exit 1
}

# Create backup of registry keys
function Backup-Registry {
    Write-ColorOutput "Creating registry backup..." "Yellow"
    
    $backupDir = "$env:USERPROFILE\Desktop\ProfileFix_Backup"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }
    
    $backupFile = "$backupDir\ProfileList_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    reg export "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" $backupFile /y
    
    Write-ColorOutput "Registry backup saved to: $backupFile" "Green"
}

# Fix 1: Reset the User Profile Service
function Reset-ProfileService {
    Write-ColorOutput "Resetting User Profile Service..." "Yellow"
    
    try {
        # Stop the User Profile Service
        Stop-Service -Name "ProfSvc" -Force -ErrorAction Stop
        Write-ColorOutput "User Profile Service stopped successfully." "Green"
        
        # Start the User Profile Service
        Start-Service -Name "ProfSvc" -ErrorAction Stop
        Write-ColorOutput "User Profile Service restarted successfully." "Green"
    }
    catch {
        Write-ColorOutput "ERROR resetting User Profile Service: $_" "Red"
    }
}

# Fix 2: Repair corrupted profile registry entries
function Repair-ProfileRegistry {
    Write-ColorOutput "Repairing profile registry entries..." "Yellow"
    
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $profileKeys = Get-ChildItem -Path $profileListPath
    
    foreach ($key in $profileKeys) {
        $sid = Split-Path -Path $key.PSPath -Leaf
        
        # Skip non-user SIDs
        if ($sid -notmatch 'S-1-5-21') { continue }
        
        # Fix 2.1: Fix duplicate .bak profile keys
        if ($sid.EndsWith(".bak")) {
            $originalSid = $sid -replace "\.bak$", ""
            $originalExists = Test-Path "$profileListPath\$originalSid"
            
            if ($originalExists) {
                Write-ColorOutput "Found duplicate profile key with .bak extension: $sid" "Yellow"
                
                # Prompt user about which profile to keep
                Write-ColorOutput "Would you like to remove the .bak profile key? This might resolve profile loading issues. (Y/N)" "Cyan"
                $response = Read-Host
                
                if ($response.ToUpper() -eq "Y") {
                    Remove-Item -Path "$profileListPath\$sid" -Force -Recurse
                    Write-ColorOutput "Removed duplicate .bak profile key: $sid" "Green"
                }
            }
        }
        
        # Fix 2.2: Check and fix ProfileImagePath
        $profilePath = (Get-ItemProperty -Path $key.PSPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue).ProfileImagePath
        
        if ($profilePath -and -not (Test-Path $profilePath)) {
            Write-ColorOutput "Profile path does not exist: $profilePath for SID $sid" "Yellow"
            
            # Check if there's a similarly named folder in C:\Users
            $username = Split-Path -Path $profilePath -Leaf
            $possiblePath = "C:\Users\$username"
            
            if (Test-Path $possiblePath) {
                Write-ColorOutput "Found possible correct path: $possiblePath" "Green"
                
                # Set the correct path
                Set-ItemProperty -Path $key.PSPath -Name "ProfileImagePath" -Value $possiblePath
                Write-ColorOutput "Fixed ProfileImagePath for SID $sid" "Green"
            }
        }
        
        # Fix 2.3: Fix State flag if incorrect
        $stateValue = (Get-ItemProperty -Path $key.PSPath -Name "State" -ErrorAction SilentlyContinue).State
        if ($null -ne $stateValue -and $stateValue -ne 0) {
            Write-ColorOutput "Profile has non-zero state value ($stateValue) for SID $sid" "Yellow"
            Set-ItemProperty -Path $key.PSPath -Name "State" -Value 0
            Write-ColorOutput "Reset State value to 0 for SID $sid" "Green"
        }
    }
}

# Fix 3: Repair file system permissions
function Repair-ProfilePermissions {
    Write-ColorOutput "Repairing profile folder permissions..." "Yellow"
    
    $userFolders = Get-ChildItem -Path "C:\Users" -Directory | Where-Object { $_.Name -notmatch "Public|Default|defaultuser0|All Users" }
    
    foreach ($folder in $userFolders) {
        $folderName = $folder.Name
        Write-ColorOutput "Checking permissions for profile folder: $folderName..." "Yellow"
        
        # First try to find the SID from the registry that maps to this folder name
        $foundSid = $null
        $profileList = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | 
                       Where-Object { $_.PSChildName -match 'S-1-5-21' }
                       
        foreach ($profile in $profileList) {
            $profilePath = (Get-ItemProperty -Path $profile.PSPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue).ProfileImagePath
            if ($profilePath -and $profilePath -eq "C:\Users\$folderName") {
                $foundSid = $profile.PSChildName
                break
            }
        }
        
        # If we couldn't find it in registry, try to translate the folder name to a SID as fallback
        if (-not $foundSid) {
            try {
                $ntAccount = New-Object System.Security.Principal.NTAccount("$folderName")
                $foundSid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
            catch {
                Write-ColorOutput "Could not determine SID for folder $folderName, using folder name as username..." "Yellow"
                $foundSid = $null
            }
        }
        
        # Use the actual folder name for operations regardless of SID resolution
        $folderName = $folder.Name
        $folderPath = $folder.FullName
        
        # Check if profile exists in registry if we found a SID
        $profileInRegistry = $false
        if ($foundSid) {
            $profileInRegistry = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$foundSid"
            if (-not $profileInRegistry) {
                $profileInRegistry = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$foundSid.bak"
            }
        }
        
        # Proceed with repair even if not in registry - using folder name as the actual username
        try {
            # Reset permissions on the profile folder
            Write-ColorOutput "Repairing permissions for folder: $folderPath" "Yellow"
            
            # Take ownership
            $takeown = Start-Process -FilePath "takeown.exe" -ArgumentList "/f `"$folderPath`" /r /d y" -NoNewWindow -Wait -PassThru
            
            # Grant the user full control - using actual folder name
            $icacls = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$folderPath`" /grant `"$folderName`":(F) /t /c /q" -NoNewWindow -Wait -PassThru
                
                if ($takeown.ExitCode -eq 0 -and $icacls.ExitCode -eq 0) {
                    Write-ColorOutput "Successfully repaired permissions for $folderName profile" "Green"
                }
                else {
                    Write-ColorOutput "Error resetting permissions for $folderName profile" "Red"
                }
                
                # Fix AppData folder specifically
                $appDataPath = "$folderPath\AppData"
                if (Test-Path $appDataPath) {
                    $takeownAppData = Start-Process -FilePath "takeown.exe" -ArgumentList "/f `"$appDataPath`" /r /d y" -NoNewWindow -Wait -PassThru
                    $icaclsAppData = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$appDataPath`" /grant `"$folderName`":(F) /t /c /q" -NoNewWindow -Wait -PassThru
                    
                    if ($takeownAppData.ExitCode -eq 0 -and $icaclsAppData.ExitCode -eq 0) {
                        Write-ColorOutput "Successfully repaired permissions for $folderName AppData folder" "Green"
                    }
                }
            }
            catch {
                Write-ColorOutput "ERROR repairing permissions: $_" "Red"
        }
        else {
            # Just inform but continue anyway
            Write-ColorOutput "No registry entry found for $folderName profile folder, proceeding with repair anyway." "Yellow"
        }
    }
    
    Write-ColorOutput "Profile permission repairs completed." "Green"
}

# Fix 4: Rebuild corrupted Group Policy settings
function Repair-GroupPolicy {
    Write-ColorOutput "Resetting Group Policy settings related to profiles..." "Yellow"
    
    try {
        # Force a Group Policy update
        gpupdate /force
        
        # Reset User Profile service-related Group Policy settings
        $gpoPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
        
        # Check if path exists
        if (Test-Path $gpoPath) {
            # Remove any problematic profile settings
            $keysToCheck = @(
                "DeleteRoamingCache",
                "EnableProfileQuota",
                "ProfileQuotaWarning",
                "ProfileQuotaLimit",
                "ProfileErrorAction",
                "ProfileDlgTimeOut"
            )
            
            foreach ($key in $keysToCheck) {
                if (Get-ItemProperty -Path $gpoPath -Name $key -ErrorAction SilentlyContinue) {
                    Remove-ItemProperty -Path $gpoPath -Name $key -Force
                    Write-ColorOutput "Removed potentially problematic GPO setting: $key" "Green"
                }
            }
        }
        
        Write-ColorOutput "Group Policy settings reset complete." "Green"
    }
    catch {
        Write-ColorOutput "ERROR resetting Group Policy settings: $_" "Red"
    }
}

# Fix 5: Repair system files
function Repair-SystemFiles {
    Write-ColorOutput "Running System File Checker to repair system files..." "Yellow"
    Write-ColorOutput "This may take several minutes..." "Yellow"
    
    # Run SFC
    $sfc = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -Wait -PassThru
    
    if ($sfc.ExitCode -eq 0) {
        Write-ColorOutput "System File Checker completed successfully." "Green"
    }
    else {
        Write-ColorOutput "System File Checker encountered issues. Exit code: $($sfc.ExitCode)" "Yellow"
    }
    
    # Run DISM
    Write-ColorOutput "Running DISM to repair Windows image..." "Yellow"
    Write-ColorOutput "This may take several minutes..." "Yellow"
    
    $dism = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -NoNewWindow -Wait -PassThru
    
    if ($dism.ExitCode -eq 0) {
        Write-ColorOutput "DISM repair completed successfully." "Green"
    }
    else {
        Write-ColorOutput "DISM repair encountered issues. Exit code: $($dism.ExitCode)" "Yellow"
    }
}

# Fix 6: Reset Windows Credential Manager
function Reset-CredentialManager {
    Write-ColorOutput "Resetting Windows Credential Manager..." "Yellow"
    
    try {
        # Stop the Credential Manager service
        Stop-Service -Name "VaultSvc" -Force -ErrorAction Stop
        
        # Clear credentials cache
        $credPath = "$env:LOCALAPPDATA\Microsoft\Credentials"
        if (Test-Path $credPath) {
            Write-ColorOutput "Backing up credentials to Desktop..." "Yellow"
            $credBackupPath = "$env:USERPROFILE\Desktop\CredentialBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            Copy-Item -Path $credPath -Destination $credBackupPath -Recurse
            
            # Clear credentials
            Remove-Item -Path "$credPath\*" -Force -Recurse
            Write-ColorOutput "Credentials cache cleared." "Green"
        }
        
        # Start the service again
        Start-Service -Name "VaultSvc" -ErrorAction Stop
        Write-ColorOutput "Credential Manager reset successfully." "Green"
    }
    catch {
        Write-ColorOutput "ERROR resetting Credential Manager: $_" "Red"
    }
}

# Fix 7: Clean up temporary profiles
function Remove-TemporaryProfiles {
    Write-ColorOutput "Cleaning up temporary profiles..." "Yellow"
    
    $tempProfiles = Get-ChildItem -Path "C:\Users" -Filter "TEMP*" -Directory -ErrorAction SilentlyContinue
    
    if ($tempProfiles) {
        foreach ($profile in $tempProfiles) {
            try {
                Remove-Item -Path $profile.FullName -Force -Recurse
                Write-ColorOutput "Removed temporary profile: $($profile.FullName)" "Green"
            }
            catch {
                Write-ColorOutput "ERROR removing temporary profile $($profile.FullName): $_" "Red"
            }
        }
    }
    else {
        Write-ColorOutput "No temporary profiles found." "Green"
    }
}

# Fix 8: Fix specific 23H2 issues - Registry tweaks
function Set-23H2Fixes {
    Write-ColorOutput "Setting specific Windows 11 23H2 fixes..." "Yellow"
    
    try {
        # Fix 1: Adjust User Profile Service start type
        Set-Service -Name "ProfSvc" -StartupType Automatic
        
        # Fix 2: Add registry optimization for User Profile Service
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\ProfSvc\Parameters"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Set optimized values
        Set-ItemProperty -Path $regPath -Name "ServiceDll" -Value "%SystemRoot%\System32\profsvc.dll" -Type ExpandString
        Set-ItemProperty -Path $regPath -Name "ServiceMain" -Value "UserProfileServiceMain" -Type String
        Set-ItemProperty -Path $regPath -Name "ServiceDllUnloadOnStop" -Value 1 -Type DWord
        
        # Fix 3: Fix User Profile event channel
        $eventCmd = "wevtutil set-log Microsoft-Windows-User Profile Service/Operational /enabled:true"
        Invoke-Expression $eventCmd
        
        Write-ColorOutput "Windows 11 23H2 specific fixes applied." "Green"
    }
    catch {
        Write-ColorOutput "ERROR applying 23H2 fixes: $_" "Red"
    }
}

# Fix 9: Clear profile cache on domain controllers (if applicable)
function Clear-DomainProfileCache {
    Write-ColorOutput "Checking for domain environment..." "Yellow"
    
    $computerSystem = Get-CimInstance Win32_ComputerSystem
    if ($computerSystem.PartOfDomain) {
        Write-ColorOutput "Machine is domain-joined to: $($computerSystem.Domain)" "Green"
        
        # Prompt for domain admin credentials for remote operations
        Write-ColorOutput "Would you like to clear profile cache on domain controllers? (Only select Y if you have domain admin privileges) (Y/N)" "Cyan"
        $response = Read-Host
        
        if ($response.ToUpper() -eq "Y") {
            $domainCred = Get-Credential -Message "Enter domain admin credentials" -UserName "$($computerSystem.Domain)\administrator"
            
            if ($domainCred) {
                try {
                    # Get domain controllers
                    $dcs = Get-ADDomainController -Filter * -Credential $domainCred -ErrorAction Stop |
                           Select-Object -ExpandProperty Name
                    
                    foreach ($dc in $dcs) {
                        Write-ColorOutput "Connecting to domain controller: $dc" "Yellow"
                        
                        # Clear profile cache
                        $session = New-PSSession -ComputerName $dc -Credential $domainCred -ErrorAction Stop
                        
                        Invoke-Command -Session $session -ScriptBlock {
                            # Restart Netlogon service (refreshes profile cache)
                            Restart-Service -Name Netlogon -Force
                            Write-Output "Netlogon service restarted on $env:COMPUTERNAME"
                        }
                        
                        Remove-PSSession $session
                        Write-ColorOutput "Profile cache cleared on domain controller: $dc" "Green"
                    }
                }
                catch {
                    Write-ColorOutput "ERROR clearing domain profile cache: $_" "Red"
                    Write-ColorOutput "You may need to contact your domain administrator to clear profile caches on the domain controllers." "Yellow"
                }
            }
        }
    }
    else {
        Write-ColorOutput "Machine is not domain-joined. Skipping domain profile cache clearing." "Green"
    }
}

# Fix 10: Update user profile caching behavior
function Update-ProfileCaching {
    Write-ColorOutput "Updating profile caching behavior..." "Yellow"
    
    try {
        # Enable profile unloading
        $unloadPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        Set-ItemProperty -Path $unloadPath -Name "AutoRestartShell" -Value 1 -Type DWord
        
        # Set cached logons count for domain-joined machines
        $computerSystem = Get-CimInstance Win32_ComputerSystem
        if ($computerSystem.PartOfDomain) {
            Set-ItemProperty -Path $unloadPath -Name "CachedLogonsCount" -Value 10 -Type String
            Write-ColorOutput "Updated CachedLogonsCount to 10" "Green"
        }
        
        Write-ColorOutput "Profile caching behavior updated." "Green"
    }
    catch {
        Write-ColorOutput "ERROR updating profile caching: $_" "Red"
    }
}

# Main script execution
try {
    # Create registry backup before making changes
    Backup-Registry
    
    # Start with easier fixes, then progress to more invasive ones
    Reset-ProfileService
    Repair-ProfileRegistry
    Reset-CredentialManager
    Remove-TemporaryProfiles
    Update-ProfileCaching
    Set-23H2Fixes
    
    # Ask user if they want to continue with more invasive fixes
    Write-ColorOutput "Initial fixes complete. Would you like to continue with more comprehensive fixes?" "Cyan"
    Write-ColorOutput "These may take longer and require a system restart afterwards. (Y/N)" "Cyan"
    $response = Read-Host
    
    if ($response.ToUpper() -eq "Y") {
        Repair-ProfilePermissions
        Repair-GroupPolicy
        Clear-DomainProfileCache
        Repair-SystemFiles
        
        Write-ColorOutput "All fixes have been applied." "Green"
        Write-ColorOutput "It is STRONGLY recommended to restart your computer now." "Yellow"
        Write-ColorOutput "Would you like to restart now? (Y/N)" "Cyan"
        $restart = Read-Host
        
        if ($restart.ToUpper() -eq "Y") {
            Restart-Computer -Force
        }
    }
    else {
        Write-ColorOutput "Basic fixes have been applied. Please restart your computer to apply changes." "Yellow"
    }
}
catch {
    Write-ColorOutput "ERROR: Script execution failed: $_" "Red"
    Write-ColorOutput "Please check the error message and try again." "Red"
}