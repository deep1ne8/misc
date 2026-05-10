# 3CX Provisioning URL Generator with Interactive Prompts
# This script generates provisioning URLs for 3CX users by connecting to the 3CX API

param(
    [switch]$Silent,
    [string]$ConfigFile
)

# Function to securely prompt for password
function Get-SecureInput {
    param(
        [string]$Prompt,
        [switch]$AsSecureString
    )
    
    if ($AsSecureString) {
        $secureInput = Read-Host -Prompt $Prompt -AsSecureString
        return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureInput))
    } else {
        return Read-Host -Prompt $Prompt
    }
}

# Function to validate URL format
function Test-UrlFormat {
    param([string]$Url)
    
    if ($Url -match '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(:\d+)?$') {
        return $true
    }
    return $false
}

# Function to load configuration from file
function Import-Configuration {
    param([string]$Path)
    
    if (Test-Path $Path) {
        try {
            $config = Get-Content $Path | ConvertFrom-Json
            return $config
        } catch {
            Write-Warning "Failed to load configuration file: $_"
            return $null
        }
    }
    return $null
}

# Function to save configuration to file
function Export-Configuration {
    param(
        [hashtable]$Config,
        [string]$Path
    )
    
    try {
        $Config | ConvertTo-Json | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Configuration saved to: $Path" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to save configuration: $_"
    }
}

# Main script execution
try {
    Write-Host "=== 3CX Provisioning URL Generator ===" -ForegroundColor Cyan
    Write-Host ""

    # Load configuration if specified
    $config = $null
    if ($ConfigFile -and (Test-Path $ConfigFile)) {
        $config = Import-Configuration -Path $ConfigFile
        if ($config) {
            Write-Host "Configuration loaded from: $ConfigFile" -ForegroundColor Green
        }
    }

    # Prompt for input parameters
    if (-not $Silent) {
        Write-Host "Enter the following information:" -ForegroundColor Yellow
        Write-Host ""
    }

    # 3CX Server URL
    if ($config -and $config.tcxurl) {
        $tcxurl = $config.tcxurl
        Write-Host "Using saved 3CX URL: $tcxurl" -ForegroundColor Gray
    } else {
        do {
            $tcxurl = Read-Host "3CX Server URL (e.g., mypbx.mydomain.de or mypbx.mydomain.de:8443)"
            if (-not (Test-UrlFormat $tcxurl)) {
                Write-Host "Invalid URL format. Please enter a valid domain or IP address." -ForegroundColor Red
            }
        } while (-not (Test-UrlFormat $tcxurl))
    }

    # Username
    if ($config -and $config.username) {
        $username = $config.username
        Write-Host "Using saved username: $username" -ForegroundColor Gray
    } else {
        $username = Read-Host "3CX Username (extension number or email)"
    }

    # Password (always prompt for security)
    $password = Get-SecureInput -Prompt "3CX Password" -AsSecureString

    # Security Code (optional)
    if ($config -and $config.securitycode) {
        $securitycode = $config.securitycode
    } else {
        $securitycode = Read-Host "Security Code (optional, press Enter to skip)"
    }

    # Ask to save configuration
    if (-not $ConfigFile -and -not $Silent) {
        $saveConfig = Read-Host "Save configuration for future use? (y/N)"
        if ($saveConfig -eq 'y' -or $saveConfig -eq 'Y') {
            $configPath = Read-Host "Enter path to save configuration file (default: .\3cx-config.json)"
            if ([string]::IsNullOrWhiteSpace($configPath)) {
                $configPath = ".\3cx-config.json"
            }
            
            $configToSave = @{
                tcxurl = $tcxurl
                username = $username
                securitycode = $securitycode
            }
            Export-Configuration -Config $configToSave -Path $configPath
        }
    }

    Write-Host ""
    Write-Host "Connecting to 3CX server..." -ForegroundColor Yellow

    # Prepare authentication body
    $jsonauthbody = @{
        Username = $username
        Password = $password
        SecurityCode = $securitycode
    } | ConvertTo-Json

    # Authenticate to 3CX
    $loginUri = "https://$tcxurl/webclient/api/Login/GetAccessToken"
    Write-Host "Authenticating with: $loginUri" -ForegroundColor Gray
    
    $LoginResponse = Invoke-RestMethod -Uri $loginUri -Body $jsonauthbody -Method POST -ContentType "application/json" -ErrorAction Stop
    
    if (-not $LoginResponse.Token.access_token) {
        throw "Authentication failed - no access token received"
    }

    $headers = @{Authorization = "Bearer $($LoginResponse.Token.access_token)"}
    Write-Host "Authentication successful!" -ForegroundColor Green

    # Get web root parameter
    Write-Host "Retrieving web root configuration..." -ForegroundColor Yellow
    $webroot = ((Invoke-RestMethod -uri "https://$tcxurl/xapi/v1/Parameters?`$search=`"WEB_ROOT_EXT_SEC`"" -headers $headers -Method GET -ContentType "application/json").Value).Value
    
    # Get provisioning directory
    Write-Host "Retrieving provisioning directory..." -ForegroundColor Yellow
    $provdir = ((Invoke-RestMethod -uri "https://$tcxurl/xapi/v1/Parameters?`$search=`"PROVISIONING_FOLDER`"" -headers $headers -Method GET -ContentType "application/json").Value | Where-Object {$_.name -eq "PROVISIONING_FOLDER"}).Value
    
    # Get user numbers
    Write-Host "Retrieving user list..." -ForegroundColor Yellow
    $usernrs = ((Invoke-RestMethod -uri "https://$tcxurl/xapi/v1/Users" -headers $headers -Method GET -ContentType "application/json").Value).Number | Sort-Object
    
    Write-Host ""
    Write-Host "=== Provisioning URLs ===" -ForegroundColor Cyan
    Write-Host "Total users found: $($usernrs.Count)" -ForegroundColor Green
    Write-Host ""

    # Generate provisioning URLs for each user
    $results = @()
    $counter = 0
    
    foreach ($usernr in $usernrs) {
        $counter++
        Write-Progress -Activity "Processing Users" -Status "Processing user $usernr" -PercentComplete (($counter / $usernrs.Count) * 100)
        
        try {
            $provfilestr = ((Invoke-RestMethod -uri "https://$tcxurl/xapi/v1/DNProperties/Pbx.GetPropertiesByDn(dnnumber='$($usernr)')?`$search=`"PROVFILE`"" -headers $headers -Method GET -ContentType "application/json").Value).Value
            
            if ($provfilestr) {
                $provisioningUrl = "tcx+app:$($webroot)provisioning/$provdir/TcxProvFiles/$provfilestr"
                $results += [PSCustomObject]@{
                    UserNumber = $usernr
                    ProvisioningFile = $provfilestr
                    ProvisioningURL = $provisioningUrl
                }
                Write-Host "User $usernr`: $provisioningUrl" -ForegroundColor White
            } else {
                Write-Host "User $usernr`: No provisioning file found" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "User $usernr`: Error retrieving provisioning file - $_" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Processing Users" -Completed
    Write-Host ""
    Write-Host "Processing completed!" -ForegroundColor Green
    Write-Host "Successfully processed $($results.Count) users with provisioning URLs" -ForegroundColor Green

    # Option to export results
    if (-not $Silent) {
        $exportResults = Read-Host "Export results to CSV? (y/N)"
        if ($exportResults -eq 'y' -or $exportResults -eq 'Y') {
            $csvPath = Read-Host "Enter CSV file path (default: .\3cx-provisioning-urls.csv)"
            if ([string]::IsNullOrWhiteSpace($csvPath)) {
                $csvPath = ".\3cx-provisioning-urls.csv"
            }
            
            $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Results exported to: $csvPath" -ForegroundColor Green
        }
    }

} catch {
    Write-Host ""
    Write-Host "Error occurred: $_" -ForegroundColor Red
    Write-Host "Stack trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}

# Usage examples:
# .\script.ps1                                    # Interactive mode
# .\script.ps1 -ConfigFile ".\my-config.json"    # Load from config file
# .\script.ps1 -Silent                           # Minimal output (requires config file)
