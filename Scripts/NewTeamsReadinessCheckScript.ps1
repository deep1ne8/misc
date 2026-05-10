################################################################################
# MIT License
#
# Copyright (c) 2024 Microsoft and Contributors
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# 
# Filename: NewTeamsReadinessCheckScript.ps1
# Version: 1.6.0.0
# Update Notes: 1.6.0.0: Added compatability mode check (thanks to Ste, Chris)
# Description: Script to validate readiness for installation of the New Teams client
#################################################################################

$script:ScriptName  = "Microsoft new Teams Installation Readiness Check"
$script:Version     = "1.6.0.0"
$ver = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'
$script:osVersion = $ver.DisplayVersion
if($script:osVersion -eq "")    {
        $script:osVersion = $ver.ReleaseId
}
$script:osBuild = (Get-WmiObject -Class Win32_OperatingSystem).Version
$script:osUBR= [int]$ver.UBR
$script:osFullBuild = [version]"$script:osBuild.$script:osUBR"
$script:osProductName = $ver.ProductName

$script:profilePaths = @(
    "$env:APPDATA",
    "$env:APPDATA\Microsoft",
    "$env:APPDATA\Microsoft\Crypto",
    "$env:APPDATA\Microsoft\Internet Explorer",
    "$env:APPDATA\Microsoft\Internet Explorer\UserData",
    "$env:APPDATA\Microsoft\Internet Explorer\UserData\Low",
    "$env:APPDATA\Microsoft\Spelling",
    "$env:APPDATA\Microsoft\SystemCertificates",
    "$env:APPDATA\Microsoft\Windows",
    "$env:APPDATA\Microsoft\Windows\Libraries",
    "$env:APPDATA\Microsoft\Windows\Recent",
    "$env:LOCALAPPDATA",
    "$env:LOCALAPPDATA\Microsoft",
    "$env:LOCALAPPDATA\Microsoft\Windows",
    "$env:LOCALAPPDATA\Microsoft\Windows\Explorer",
    "$env:LOCALAPPDATA\Microsoft\Windows\History",
    "$env:LOCALAPPDATA\Microsoft\Windows\History\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\History\Low\History.IE5",
    "$env:LOCALAPPDATA\Microsoft\Windows\IECompatCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\IECompatCache\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\IECompatUaCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\IECompatUaCache\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies\DNTException",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies\DNTException\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies\PrivacIE",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCookies\PrivacIE\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\PPBCompatCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\PPBCompatCache\Low",
    "$env:LOCALAPPDATA\Microsoft\Windows\PPBCompatUaCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\PPBCompatUaCache\Low",
    "$env:LOCALAPPDATA\Microsoft\WindowsApps",
    "$env:LOCALAPPDATA\Packages",
    "$env:LOCALAPPDATA\Publishers",
    "$env:LOCALAPPDATA\Publishers\8wekyb3d8bbwe",
    "$env:LOCALAPPDATA\Temp",
    "$env:USERPROFILE\AppData\LocalLow",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer\DOMStore",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer\EdpDomStore",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer\EmieSiteList",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer\EmieUserList",
    "$env:USERPROFILE\AppData\LocalLow\Microsoft\Internet Explorer\IEFlipAheadCache"
)

$script:Errors = @()
$script:Out = @()

function Run()
{
    isTeamsInCompatibilityMode
    if (isNewTeamsNotInstalled)
    {
        ValidateOSVersion
        ValidateAppXPolicies
        ValidateWebView2Installed
        ValidateT1Version
        isDOMaxForegroundDownloadBandwidthConfigured
        isDODownloadModeConfigured
    }
    
    ValidateShellFolders
    ValidatePathTypes
    ValidateReparsePoints
    ValidateUserAccess
    ValidateSystemPerms
    ValidateSysAppIdPerms
    OutputExit
}

function isTeamsInCompatibilityMode()
{
    # Check for app compatibility entries for Teams in the user hive
    Get-Item -Path Registry::'HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers' | Select-Object -ExpandProperty Property | ForEach-Object { if ($_ -match 'Teams') { WriteError "User level Teams compatibility mode entry detected. Please run Teams in normal mode." }}

    # Check for app compatibility entries for Teams in the local machine hive
    Get-Item -Path Registry::'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers' | Select-Object -ExpandProperty Property | ForEach-Object { if ($_ -match 'Teams') { WriteError "Computer level Teams compatibility mode entry detected. Please run Teams in normal mode." }}

}

function isNewTeamsNotInstalled()
{
    if ((Get-appxpackage -name MSTeams).PackageFamilyName -eq "MSTeams_8wekyb3d8bbwe")
    {
        return $false
    }
        
    return $true
}

function WriteError([string]$line)
{
    $script:Errors += $line +"`n"
}

function WriteOut([string]$line)
{
    $script:Out += $line + "`n"
}

function OutputExit
{
    if ($script:errors.Count -gt 0)
    {
        try
        {
            #to know how many machines have issues
            Invoke-WebRequest -Uri "https://teams.microsoft.com/appdiag/newteamsreadinesscheckerror"
        } catch {
        }

        Write-Host "Microsoft new Teams pre-install validation Summary.`n `n Following issues need to be fixed for the new Teams installation to succeed and launch without any issues:`n $script:errors `n`n Following passed and no action needed:`n $script:Out "
        Exit 1
    }
    else
    {
          try
        {
            #to know how many machines don't have issues
            Invoke-WebRequest -Uri "https://teams.microsoft.com/appdiag/newteamsreadinesschecksuccess"
        } catch {
        }
        Write-Host "We did not find any issues. You should be ready for the new Teams Installation. Following checks were performed:`n $script:Out"
        Exit 0
    }
}

# This function validates the regkeys at User shell folder path 
function ValidateShellFolders
{
    $badPaths = 0

    $expectedRegkeyValues = @{
        "Cookies" = "$env:USERPROFILE\AppData\Local\Microsoft\Windows\INetCookies"
        "Cache" = "$env:USERPROFILE\AppData\Local\Microsoft\Windows\INetCache"
    }

    $actualRegkeyValues = @{}
    $shellFolderPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

    # Retrieve the current values of the shell folders
    try {
        $itemProps = Get-ItemProperty -Path $shellFolderPath
        foreach ($key in $expectedRegkeyValues.Keys) {
            $actualRegkeyValues[$key] = $itemProps.$key
        }
    } catch {
        WriteError "Unable to access registry path: $shellFolderPath. Please make sure the path exists."
        return
    }

    # Validate and correct the paths if necessary
    foreach ($regkeyValue in $expectedRegkeyValues.Keys) {
        $regkeyPathValue = $actualRegkeyValues[$regkeyValue]
        if (listReparsePaths($regkeyPathValue) -neq "") {
                WriteError "Value of regkey $regkeyValue at $shellFolderPath contains reparse point(s). Set $regkeyValue at $shellFolderPath to $expectedRegkeyValues[$regkeyValue]"
                $badPaths++
        } else {
            WriteOut "$regkeyValue path is correct $regkeyPathValue."
        }
    }

    if ($badPaths -eq 0) {
        WriteOut "No issues found with the Reg key values Cookies & Cache."
    }
}

# This function validates the reparse points for App data folders
function ValidateReparsePoints
{
    $badPaths = ""

    foreach ($path in $script:profilePaths)
    {
        $badPaths += listReparsePaths($path)
    }

    if ([string]::IsNullOrEmpty($badPaths))
    {
        WriteOut "No reparse points found."
    } else
    {
        WriteError "new Teams installation fails when the installer hits a reparse path. Please fix the following Reparse paths:`n$badPaths"
    }
}

# This function validates the right user permission for App data folders
function ValidateUserAccess
{
    $badPaths = ""
    foreach ($path in $script:profilePaths)
    {
        $left = $path
        for($i=0;$i -lt 10; $i++)
        {
            if ([string]::IsNullOrEmpty($left))
            {
                break;
            };
            try
            {
                if (Test-Path -Path $left)
                {
                   $items = Get-ChildItem $left -ErrorAction SilentlyContinue -ErrorVariable GCIErrors
                   if($GCIErrors.Count -gt 0)
                   {
                     $badPaths += "$left`n"
                   }
                }
            }
            catch
            {
                WriteError "Unexpected Error during file access check."     
            }
            $left=Split-Path $left
        }
    }
    
    if ([string]::IsNullOrEmpty($badPaths))
    {
        WriteOut "All paths are accessible. No issues found."
    } else
    {
        WriteError "We couldn't access the following paths. Please make the following paths readable for the new Teams installation to pass:`n$badPaths"
    }
}

# This function validates the SYSTEM and Administrator group permission on App data folders
function ValidateSystemPerms
{
    $badPaths = ""
    foreach ($path in $script:profilePaths)
    {
        try{
            if (Test-Path -Path $path)
            {
                try{                
                    $systemPerms = (Get-Acl $path).Access | where {$_.IdentityReference -eq "NT AUTHORITY\SYSTEM"}
                    $systemFullControlAllow = $systemPerms | where {$_.FileSystemRights -eq "FullControl" -and $_.AccessControlType -eq "Allow"}
                    $systemFullControlDeny = $systemPerms | where {$_.FileSystemRights -eq "FullControl" -and $_.AccessControlType -eq "Deny"}
                    if($systemFullControlAllow.Count -eq 0 -or $systemFullControlDeny.Count -gt 0)
                    {
                        $badPaths += "$path is missing permissions for the SYSTEM account."
                    } 
                }catch
                {
                    $badPaths += "$path is missing permissions for the SYSTEM account."
                }
                try{
                    $adminPerms = (Get-Acl $path).Access | where {$_.IdentityReference -eq "BUILTIN\Administrators"}
                    $adminFullControlAllow = $adminPerms | where {$_.FileSystemRights -eq "FullControl" -and $_.AccessControlType -eq "Allow"}
                    $adminFullControlDeny = $adminPerms | where {$_.FileSystemRights -eq "FullControl" -and $_.AccessControlType -eq "Deny"}
                    if($adminFullControlAllow.Count -eq 0 -or $adminFullControlDeny.Count -gt 0)
                    {
                        $badPaths += "$path is missing permissions for the Administrators group.`n"
                    }
                }catch
                {
                    $badPaths += "$path is missing permissions for the Administrators group.`n"
                }
            }
        }catch{
            $badPaths += "User is not able to access $path.`n"
        }

    }
    
    if ([string]::IsNullOrEmpty($badPaths))
    {
        WriteOut "Ran System and Administrator permission checks. No issues found."
    } else
    {
        WriteError "SYSTEM or Administrator group permissions are required for the new Teams installation. Please fix the following paths:`n$badPaths"
    }
}

# This function validates the right permission on Teams installation directory
function ValidateSysAppIdPerms
{
    $apps = Get-AppxPackage MSTeams 
    foreach($app in $apps)
    {
        try{
            $perms = (Get-Acl $app.InstallLocation).sddl -split "\(" | ?{$_ -match "WIN:/\/\SYSAPPID"}
            if($perms.Length -gt 0)
            {
                WriteOut "$($app.InstallLocation) has the correct SYSAPPID permissions assigned."
            }
            else
            {
                WriteError "$($app.InstallLocation) is missing SYSAPPID permissions. This is required for successful launch of the new Teams."
            }
        }catch{
            WriteError "$($app.InstallLocation) is missing SYSAPPID permissions. This is required for successful launch of the new Teams."
        }
    }
}

# This function validates whether App data paths are intact or not
function ValidatePathTypes
{
    $badFolders = ""
    $badPaths = ""
    foreach ($path in $script:profilePaths)
    {
        if (Test-Path -Path $path)
        {
            if(!(Test-Path -Path $path -PathType Container))
            {
                $badFolders = $badFolders + $path + "`n"
            }
        }
        Else
        {
            $badPaths = $badPaths + $path + "`n"
        }
    }

    if ([string]::IsNullOrEmpty($badPaths) -and [string]::IsNullOrEmpty($badFolders))
    {
        WriteOut "No issues with folders."
    }
    else
    {
        if (-Not [string]::IsNullOrEmpty($badPaths)){
            WriteError "Create empty folders at the following paths:`n$badPaths"
        }
        if (-Not [string]::IsNullOrEmpty($badFolders)){
            WriteError "Convert the following files to folders:`n$badFolders"
        }
    }
}

# This function validates the OS version
function ValidateOSVersion()
{

    $minOsVersion = [version]"10.0.19041"

    if($script:osProductName.Contains("LTSC") -or $script:osProductName.Contains("LTSB"))
    {
        WriteError "Microsoft Teams is not supported on Windows LTSC/LTSB. Please update to a non LTSC/LTSB version of Windows with a version $minOsVersion or later"
        return
    }
    if($script:osFullBuild -lt $minOsVersion)
    {
        WriteError "Your version of Windows is not supported by New Teams. Please update to $minOsVersion or later"
        return
    }
    WriteOut "No issues with the OS version."
}


# If OS is below patch version, we must have AllowAllTrustedApps keys enabled
# Else above min version but below max patch version, either have patch present or if we are below required below patch, we should have AllowAllTrustedApps keys enabled.
# If we are above maxpatchversion, we are all good
function ValidateAppXPolicies()
{
    $osPatchThresholds = @{
        "10.0.19044" = 4046 #Win 10 21H2
        "10.0.19045" = 3636 #Win 10 22H2
        "10.0.22000" = 2777 #Win 11 21H2
        "10.0.22621" = 2506 #Win 11 22H2
    }

    $minPatchVersion = [version]"10.0.19044"
    $maxPatchVersion = [version]"10.0.22621"

    if($script:osFullBuild -lt $minPatchVersion)
    {
        if(-Not (HasAllowAllTrustedAppsKeyEnabled))
        {
            WriteError "AllowAllTrustedApps is not enabled. Please enable it."
            return
        }
    }
    elseif($script:osFullBuild -le $maxPatchVersion)
    {
        $targetUBR = $osPatchThresholds[$script:osBuild]
        if($script:osUBR -lt $targetUBR)
        {
            if(-Not (HasAllowAllTrustedAppsKeyEnabled))
            {
                $recommendedVersion = [version]"$script:osBuild.$targetUBR"
                WriteError "AllowAllTrustedApps is not enabled and your version of Windows does not contain a required patch to support this.`nEither update your version of Windows to be greater than $recommendedVersion, or enable AllowAllTrustedApps."
                return
            }
        }
    }

    WriteOut "No issues related to the AppX policies."
}

# This function checks whether AllowAllTrustedApps is enabled or not
function HasAllowAllTrustedAppsKeyEnabled
{
    $hasKey = $false;
    $appXKeys = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\AppModelUnlock", "HKLM:\Software\Policies\Microsoft\Windows\Appx")
    foreach ($key in $appXKeys)
    {
        try
        {
            $value = Get-ItemPropertyValue -Path $key -Name "AllowAllTrustedApps"
            if ($value -ne 0)
            {
                $hasKey = $true
                break;
            }
        }
        catch
        {
            WriteOut "Missing AllowAllTrustedApps key at $key."
        }
    }
    return $hasKey
}

# This function validates whether WV2 is installed or not
function ValidateWebView2Installed()
{

    $webView2Keys = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
        "HKCU:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
        "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
        )

    $version = ""
    foreach($key in $webView2Keys)
    {

        if(Test-Path $key)
        {
            try
            {
                $version = Get-ItemPropertyValue -Path $key -Name "pv"
                if($version)
                {
                    WriteOut "Found WebView2 Runtime version $version."
                    return
                }
            }
            catch
            {
                #WriteOut "pv property not found in $key"
            }
        }
        else
        {
            #WriteOut "Key not found: $key"
        }
    }
    WriteError "WebView2 Runtime is not installed. Please install WebView2 from: https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/enterprise."
}

# This function checks for minimum T1 version
function ValidateT1Version()
{

    $minT1Version = [version]"1.6.00.27573"
    $t1Path = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"
    if(Test-Path $t1Path)
    {
        $props = Get-Item "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"
        $version = [version]$props.VersionInfo.FileVersion
        if($version -lt $minT1Version)
        {
            WriteError "$minT1Version is required to auto-update to new Teams. Please update your Classic Teams client version from $version to at least $minT1Version."
            return
        }
        WriteOut "No issues found. Your Teams Classic version, $version, supports auto-update."
        return
    }
    WriteOut "You do not have classic Teams."
}

# This function checks for DO mode configured on client machine
function isDODownloadModeConfigured
{
    $appXKey = @("HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization\")
    try
    {
        $value = Get-ItemPropertyValue -Path $appXKey -Name "DODownloadMode" -Force -ErrorAction SilentlyContinue
        if ($value -eq 100)
        {
            WriteError "DODownloadMode has unsupported value 100 is set, please refer https://learn.microsoft.com/en-us/microsoftteams/new-teams-deploy-using-policies?tabs=teams-admin-center#prerequisites"
            return
        }
    }
    catch
    {
        #this should have an error written out
        WriteError "Unable to confirm the value of DODownloadMode. Please refer https://learn.microsoft.com/en-us/microsoftteams/new-teams-deploy-using-policies?tabs=teams-admin-center#prerequisites and fix any issues."
        return
    }
    WriteOut "DODownloadMode reg key passed validation."
}

# This function checks whether any bandwidth restriction imposed on client machine
function isDOMaxForegroundDownloadBandwidthConfigured
{
    $appXKey = @("HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization\")
    try
    {
        $value = Get-ItemPropertyValue -Path $appXKey -Name "DOMaxForegroundDownloadBandwidth" -Force -ErrorAction SilentlyContinue
        if ($value)
        {
            WriteError "You have a bandwidth restriction set. Low bandwidth limit causes the download to timeout. If your Teams download or update fails, please increase the download bandwidth limit and try again."
            return
        }
    }
    catch
    {
                WriteError "Unable to confirm the value of DOMaxForegroundDownloadBandwidth. Please refer https://learn.microsoft.com/en-us/microsoftteams/new-teams-deploy-using-policies?tabs=teams-admin-center#prerequisites and fix any issues. Also, low bandwidth limit causes the download to timeout. If your Teams download or update fails, please increase the download bandwidth limit and try again."

        WriteError "Unable to confirm Hit exception in isDOMaxForegroundDownloadBandwidthConfigured."
        return
    }
    WriteOut "DOMaxForegroundDownloadBandwidth reg key passed validation."
}

function IsReparsePoint([string]$path)
{

    $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    if($props.Attributes -match 'ReparsePoint')
    {
        return $true
    }
    return $false
}

# changes made here
function listReparsePaths($path)
{
    $reparsePaths = ""
    $left = $path
    for($i=0;$i -lt 10; $i++)
    {
        if ([string]::IsNullOrEmpty($left))
        {
            break;
        };
        if(IsReparsePoint($left))
        {
            $reparsePaths = $reparsePaths + "$left`n"
        }
        $left=Split-Path $left
    }
    return $reparsePaths
}

Run
# SIG # Begin signature block
# MIIoUgYJKoZIhvcNAQcCoIIoQzCCKD8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBT3GPIbDOHy/O/
# 2K6Eh1VtCR0yHOCRUCI1hNheeOcux6CCDYUwggYDMIID66ADAgECAhMzAAADri01
# UchTj1UdAAAAAAOuMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwODU5WhcNMjQxMTE0MTkwODU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQD0IPymNjfDEKg+YyE6SjDvJwKW1+pieqTjAY0CnOHZ1Nj5irGjNZPMlQ4HfxXG
# yAVCZcEWE4x2sZgam872R1s0+TAelOtbqFmoW4suJHAYoTHhkznNVKpscm5fZ899
# QnReZv5WtWwbD8HAFXbPPStW2JKCqPcZ54Y6wbuWV9bKtKPImqbkMcTejTgEAj82
# 6GQc6/Th66Koka8cUIvz59e/IP04DGrh9wkq2jIFvQ8EDegw1B4KyJTIs76+hmpV
# M5SwBZjRs3liOQrierkNVo11WuujB3kBf2CbPoP9MlOyyezqkMIbTRj4OHeKlamd
# WaSFhwHLJRIQpfc8sLwOSIBBAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUhx/vdKmXhwc4WiWXbsf0I53h8T8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMTgzNjAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# AGrJYDUS7s8o0yNprGXRXuAnRcHKxSjFmW4wclcUTYsQZkhnbMwthWM6cAYb/h2W
# 5GNKtlmj/y/CThe3y/o0EH2h+jwfU/9eJ0fK1ZO/2WD0xi777qU+a7l8KjMPdwjY
# 0tk9bYEGEZfYPRHy1AGPQVuZlG4i5ymJDsMrcIcqV8pxzsw/yk/O4y/nlOjHz4oV
# APU0br5t9tgD8E08GSDi3I6H57Ftod9w26h0MlQiOr10Xqhr5iPLS7SlQwj8HW37
# ybqsmjQpKhmWul6xiXSNGGm36GarHy4Q1egYlxhlUnk3ZKSr3QtWIo1GGL03hT57
# xzjL25fKiZQX/q+II8nuG5M0Qmjvl6Egltr4hZ3e3FQRzRHfLoNPq3ELpxbWdH8t
# Nuj0j/x9Crnfwbki8n57mJKI5JVWRWTSLmbTcDDLkTZlJLg9V1BIJwXGY3i2kR9i
# 5HsADL8YlW0gMWVSlKB1eiSlK6LmFi0rVH16dde+j5T/EaQtFz6qngN7d1lvO7uk
# 6rtX+MLKG4LDRsQgBTi6sIYiKntMjoYFHMPvI/OMUip5ljtLitVbkFGfagSqmbxK
# 7rJMhC8wiTzHanBg1Rrbff1niBbnFbbV4UDmYumjs1FIpFCazk6AADXxoKCo5TsO
# zSHqr9gHgGYQC2hMyX9MGLIpowYCURx3L7kUiGbOiMwaMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGiMwghofAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAAOuLTVRyFOPVR0AAAAA
# A64wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIAFs
# y5KjUgHpmo/ZsULY0fyOOoA+9/2ef+OjCe9tX2uNMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAG0YDlPyzlFBNAnem04hMCTptdNotaW5cwm8/
# L8N961+tsvNartBZccz0khg4Z8ChUfYTzpBLZT1HV/2oV9dCXQurD3wbdfonHr8n
# VTQ3YsKDoTnbHbHTCAxSaZE/miljNXKn1tlRb0+bOAPDU5k32hhdF8e+um/esiax
# lQyzI3gXJhVo9Yi9FSgWco6+7mi6Z5ipEWkFHRkOZ5+GFnUW0U5KMpOFgfKYOkE/
# COKrR6qdix3s5Er6IAVFD5Nf8cwA52ZtVadlGQh2Yfw3T5TgkjH3WaZ2EGkWiDb2
# y3zDs04YJaimNtAPiP01BmFF30yzIL0qHq+T7zM9ZJOgcxPiZaGCF60wghepBgor
# BgEEAYI3AwMBMYIXmTCCF5UGCSqGSIb3DQEHAqCCF4YwgheCAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFaBgsqhkiG9w0BCRABBKCCAUkEggFFMIIBQQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCBRPZ68FDzT2D5Wsubmhivpk+w2t2TLFDJO
# Xux/OKodhAIGZutSBaRUGBMyMDI0MTAxMDIwNDkxNS45NzNaMASAAgH0oIHZpIHW
# MIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQL
# EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsT
# Hm5TaGllbGQgVFNTIEVTTjo0MzFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEfswggcoMIIFEKADAgECAhMzAAAB+vs7
# RNN3M8bTAAEAAAH6MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMB4XDTI0MDcyNTE4MzExMVoXDTI1MTAyMjE4MzExMVowgdMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jv
# c29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVs
# ZCBUU1MgRVNOOjQzMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# yhZVBM3PZcBfEpAf7fIIhygwYVVP64USeZbSlRR3pvJebva0LQCDW45yOrtpwIpG
# yDGX+EbCbHhS5Td4J0Ylc83ztLEbbQD7M6kqR0Xj+n82cGse/QnMH0WRZLnwggJd
# enpQ6UciM4nMYZvdQjybA4qejOe9Y073JlXv3VIbdkQH2JGyT8oB/LsvPL/kAnJ4
# 5oQIp7Sx57RPQ/0O6qayJ2SJrwcjA8auMdAnZKOixFlzoooh7SyycI7BENHTpkVK
# rRV5YelRvWNTg1pH4EC2KO2bxsBN23btMeTvZFieGIr+D8mf1lQQs0Ht/tMOVdah
# 14t7Yk+xl5P4Tw3xfAGgHsvsa6ugrxwmKTTX1kqXH5XCdw3TVeKCax6JV+ygM5i1
# NroJKwBCW11Pwi0z/ki90ZeO6XfEE9mCnJm76Qcxi3tnW/Y/3ZumKQ6X/iVIJo7L
# k0Z/pATRwAINqwdvzpdtX2hOJib4GR8is2bpKks04GurfweWPn9z6jY7GBC+js8p
# SwGewrffwgAbNKm82ZDFvqBGQQVJwIHSXpjkS+G39eyYOG2rcILBIDlzUzMFFJbN
# h5tDv3GeJ3EKvC4vNSAxtGfaG/mQhK43YjevsB72LouU78rxtNhuMXSzaHq5fFiG
# 3zcsYHaa4+w+YmMrhTEzD4SAish35BjoXP1P1Ct4Va0CAwEAAaOCAUkwggFFMB0G
# A1UdDgQWBBRjjHKbL5WV6kd06KocQHphK9U/vzAfBgNVHSMEGDAWgBSfpxVdAF5i
# XYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENB
# JTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRp
# bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsF
# AAOCAgEAuFbCorFrvodG+ZNJH3Y+Nz5QpUytQVObOyYFrgcGrxq6MUa4yLmxN4xW
# dL1kygaW5BOZ3xBlPY7Vpuf5b5eaXP7qRq61xeOrX3f64kGiSWoRi9EJawJWCzJf
# UQRThDL4zxI2pYc1wnPp7Q695bHqwZ02eaOBudh/IfEkGe0Ofj6IS3oyZsJP1yat
# cm4kBqIH6db1+weM4q46NhAfAf070zF6F+IpUHyhtMbQg5+QHfOuyBzrt67CiMJS
# KcJ3nMVyfNlnv6yvttYzLK3wS+0QwJUibLYJMI6FGcSuRxKlq6RjOhK9L3QOjh0V
# CM11rHM11ZmN0euJbbBCVfQEufOLNkG88MFCUNE10SSbM/Og/CbTko0M5wbVvQJ6
# CqLKjtHSoeoAGPeeX24f5cPYyTcKlbM6LoUdO2P5JSdI5s1JF/On6LiUT50adpRs
# tZajbYEeX/N7RvSbkn0djD3BvT2Of3Wf9gIeaQIHbv1J2O/P5QOPQiVo8+0AKm6M
# 0TKOduihhKxAt/6Yyk17Fv3RIdjT6wiL2qRIEsgOJp3fILw4mQRPu3spRfakSoQe
# 5N0e4HWFf8WW2ZL0+c83Qzh3VtEPI6Y2e2BO/eWhTYbIbHpqYDfAtAYtaYIde87Z
# ymXG3MO2wUjhL9HvSQzjoquq+OoUmvfBUcB2e5L6QCHO6qTO7WowggdxMIIFWaAD
# AgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3Nv
# ZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIy
# MjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5
# vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64
# NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhu
# je3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl
# 3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPg
# yY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I
# 5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2
# ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/
# TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy
# 16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y
# 1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6H
# XtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMB
# AAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQW
# BBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30B
# ATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMB
# Af8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBL
# oEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggr
# BgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNS
# b29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1Vffwq
# reEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27
# DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pv
# vinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9Ak
# vUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWK
# NsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2
# kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+
# c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep
# 8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+Dvk
# txW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1Zyvg
# DbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/
# 2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDVjCCAj4CAQEwggEBoYHZpIHW
# MIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQL
# EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsT
# Hm5TaGllbGQgVFNTIEVTTjo0MzFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUA94Z+bUJn+nKw
# BvII6sg0Ny7aPDaggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDANBgkqhkiG9w0BAQsFAAIFAOqyJ2kwIhgPMjAyNDEwMTAxMDE0MDFaGA8yMDI0
# MTAxMTEwMTQwMVowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA6rInaQIBADAHAgEA
# AgIg2DAHAgEAAgITwTAKAgUA6rN46QIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgor
# BgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBCwUA
# A4IBAQBP/RWXZ4ATXE8FH/DziuE8oFc4+lQkmikSbeZVLQIOKSaxlDNOHdOrxB38
# NZ2UpRUB3WHHQWa2tIlk15yKXBb5MtTlGaeCrnirs/RwUiLRpdIGZnLevPqGwnos
# aFXm456YlohVG2Huon393A3uKvoe5Eh9y2+OFZqk2TfxKz+0LGKfhjSWmGWJi5qx
# HWZXMfTXO22K+NRqp0Q2aKJMZjtJInMeWMHlWKLNrYk5YannviWsVS0C74irSe0q
# OB5Jr77jvHt6F0rhbhHQTgnf/1NvVYilqZsDzDO3e035as/q+0ZBQt7wv82QOUNy
# W60U/LJNLiBgAc/JTb9gaVjG8IZDMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAH6+ztE03czxtMAAQAAAfowDQYJYIZIAWUD
# BAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0B
# CQQxIgQgrcY9zq91Xx+QFMaxxzpRcMMi/Xu//KD43TVDytd74u4wgfoGCyqGSIb3
# DQEJEAIvMYHqMIHnMIHkMIG9BCB98n8tya8+B2jjU/dpJRIwHwHHpco5ogNStYoc
# bkOeVjCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
# +vs7RNN3M8bTAAEAAAH6MCIEIAUxL7ShMFFCzxvCbR+EvuKwf/mZiiafAedq7Tb7
# j8XFMA0GCSqGSIb3DQEBCwUABIICAGWM5gHAGFFzmIN/kA4ES0vv5bFDm8sKxedB
# Ew0mOs1Z7Krn79/Sanu0cJFYuPwJ4QuUvGVkF76l5Kj3TtYAuyXPWw6fr2nFKA+3
# 5QkzgFvL/vDFZXmHXhi5tVjuEV4o0rDNJaSweEd0MCbowE2Dy57AhSYNglqOesQh
# DxTtRc9rIHa5OHOwxxb6MQq7if/WVFjAMh1sea2UzzkTYCBxXt08IcEKjTCU1oFH
# 0WjK/2+tZTBVlefyeWUVeQUVh/SIEZbAttX/tcOl+ydIOeVlU0X9RCvMOjYMq90y
# zC0tnP8x3AMOt6+ijizWQufUYgrO+ShmyJw8o4GYTjqsqizm9vbZGko0j1usaSTj
# VFUd80KpPi5pNVOssLMQwoefAUywxB3Y1rzeRO3X+yQ4X5+fIr/M4uPS612t+5HS
# d4dU2wPoM92K2zmqqn9FlvMOcD8/qAu4scg/LMmYkwKklgOU+oiWtspxJlKs7evj
# jiBOyIRJN1f6McM9ZcGyFenIWd2XmeMyikg2/AhOPQllXBNTuAi8H5XlJeaPB5j1
# XozX+ZjkyyNTjGzzsTXosmfkfaGpMnMMPpiRFfvs1rDI51CfK1JzxkU1siHiuiLV
# GIzCCLeUMCWm/MkP9zBroTnTF+DAt/ASjDyOXL7yxsU9PeaTm0z+HOxkCHRbb/K+
# /vOyD+oC
# SIG # End signature block
