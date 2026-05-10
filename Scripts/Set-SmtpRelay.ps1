# Microsoft 365 SMTP Relay Configuration Script
# This script helps configure a server to relay emails through Microsoft 365

# Check if Exchange Online PowerShell module is installed
$moduleInstalled = Get-Module -ListAvailable -Name ExchangeOnlineManagement
if (-not $moduleInstalled) {
    Write-Host "Exchange Online Management module is not installed. Installing now..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Write-Host "Exchange Online Management module installed successfully." -ForegroundColor Green
}

# Import the Exchange Online module
Import-Module ExchangeOnlineManagement

# Function to create SMTP relay configuration
function Set-Office365SMTPRelay {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConnectorName,
        
        [Parameter(Mandatory=$true)]
        [string]$SenderDomain,
        
        [Parameter(Mandatory=$true)]
        [string]$SenderIPAddress,
        
        [Parameter(Mandatory=$true)]
        [string]$AdminUsername
    )
    
    try {
        # Connect to Exchange Online
        Write-Host "Connecting to Exchange Online. Please enter your admin credentials..." -ForegroundColor Cyan
        Connect-ExchangeOnline -UserPrincipalName $AdminUsername
        
        # Create new inbound connector for SMTP relay
        Write-Host "Creating new inbound connector for SMTP relay..." -ForegroundColor Cyan
        New-InboundConnector -Name $ConnectorName -ConnectorType OnPremises -SenderDomains $SenderDomain -SenderIPAddresses $SenderIPAddress -RequireTls $true
        
        Write-Host "SMTP relay connector created successfully!" -ForegroundColor Green
        Write-Host "Configuration Details:" -ForegroundColor Green
        Write-Host "------------------------" -ForegroundColor Green
        Write-Host "Connector Name: $ConnectorName" -ForegroundColor Green
        Write-Host "Sender Domain: $SenderDomain" -ForegroundColor Green
        Write-Host "Sender IP Address: $SenderIPAddress" -ForegroundColor Green
        Write-Host "TLS Required: Yes" -ForegroundColor Green
        
        # Create Service Account for SMTP Authentication
        $createServiceAccount = Read-Host "Do you want to create a dedicated service account for SMTP authentication? (Y/N)"
        if ($createServiceAccount -eq "Y" -or $createServiceAccount -eq "y") {
            $serviceAccountUPN = Read-Host "Enter the UPN for the service account (e.g., smtp-relay@yourdomain.com)"
            # $serviceAccountDisplayName = Read-Host "Enter the display name for the service account"
            
            # Create service account
            # New-Mailbox -Name $serviceAccountDisplayName -DisplayName $serviceAccountDisplayName -MicrosoftOnlineServicesID $serviceAccountUPN -Password (ConvertTo-SecureString -String (Read-Host "Enter password for service account" -AsSecureString) -AsPlainText -Force) -ResetPasswordOnNextLogon $false
                #Write-Host "Service Account Created Successfully!" -ForegroundColor Green

            # Configure SMTP authentication for the service account
            Set-CASMailbox -Identity $serviceAccountUPN -SmtpClientAuthenticationDisabled $false
            
            
            Write-Host "SMTP Server: smtp.office365.com" -ForegroundColor Green
            Write-Host "SMTP Port: 587" -ForegroundColor Green
            Write-Host "Authentication: Required (Username: $serviceAccountUPN)" -ForegroundColor Green
            Write-Host "Encryption: TLS" -ForegroundColor Green
        }
        
        # Disconnect from Exchange Online
        Disconnect-ExchangeOnline -Confirm:$false
    }
    catch {
        Write-Host "Error configuring SMTP relay: $_" -ForegroundColor Red
    }
}

# Function to configure the local server for SMTP relay
function Set-LocalServerForRelay {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SmtpServer,
        
        [Parameter(Mandatory=$true)]
        [string]$SmtpUsername,
        
        [Parameter(Mandatory=$true)]
        [SecureString]$SmtpPassword
    )
    
    try {
        # Create a PSCredential object for SMTP authentication
        $credential = New-Object System.Management.Automation.PSCredential($SmtpUsername, $SmtpPassword)
        
        # Test SMTP connection
        Write-Host "Testing SMTP connection to Office 365..." -ForegroundColor Cyan
        
        # Create temporary mail message for testing
        $mailParams = @{
            SmtpServer = $SmtpServer
            Port = 587
            UseSsl = $true
            From = $SmtpUsername
            To = $SmtpUsername
            Subject = "SMTP Relay Test"
            Body = "This is a test email to verify SMTP relay configuration."
            Credential = $credential
        }
        
        Send-MailMessage @mailParams
        Write-Host "Test email sent successfully! SMTP relay is working." -ForegroundColor Green
        
        # Create the SMTP configuration in Windows registry for applications
        $storeCredentials = Read-Host "Do you want to store the SMTP credentials securely for system use? (Y/N)"
        if ($storeCredentials -eq "Y" -or $storeCredentials -eq "y") {
            # Create registry keys for SMTP configuration
            $registryPath = "HKLM:\SOFTWARE\Microsoft\Office365SMTPRelay"
            
            if (!(Test-Path $registryPath)) {
                New-Item -Path $registryPath -Force | Out-Null
            }
            
            # Store SMTP server settings (encrypt password)
            New-ItemProperty -Path $registryPath -Name "SmtpServer" -Value $SmtpServer -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $registryPath -Name "SmtpPort" -Value 587 -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -Path $registryPath -Name "SmtpUsername" -Value $SmtpUsername -PropertyType String -Force | Out-Null
            
            # Convert and encrypt password
            $encryptedPassword = ConvertFrom-SecureString -SecureString $SmtpPassword
            New-ItemProperty -Path $registryPath -Name "SmtpPassword" -Value $encryptedPassword -PropertyType String -Force | Out-Null
            
            Write-Host "SMTP credentials stored securely in registry." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error configuring local server: $_" -ForegroundColor Red
    }
}

# Main script execution
Clear-Host
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "   Microsoft 365 SMTP Relay Configuration Script    " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Configure Office 365 SMTP Relay
$setupOffice365 = Read-Host "Do you want to set up the Office 365 connector? (Y/N)"
if ($setupOffice365 -eq "Y" -or $setupOffice365 -eq "y") {
    $connectorName = Read-Host "Enter a name for the connector (e.g., On-Premises-SMTP-Relay)"
    $senderDomain = Read-Host "Enter the sender domain (e.g., yourdomain.com)"
    $senderIPAddress = Read-Host "Enter the sender IP address (Party B's public IP)"
    $adminUsername = Read-Host "Enter your Microsoft 365 admin email address"
    
    Set-Office365SMTPRelay -ConnectorName $connectorName -SenderDomain $senderDomain -SenderIPAddress $senderIPAddress -AdminUsername $adminUsername
}


Write-Host ""
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "   Microsoft 365 SMTP Relay Configuration Complete  " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

# Usage Instructions
Write-Host ""
Write-Host "Usage Instructions:" -ForegroundColor Yellow
Write-Host "1. For applications, use the following SMTP settings:" -ForegroundColor Yellow
Write-Host "   - SMTP Server: smtp.office365.com" -ForegroundColor Yellow
Write-Host "   - SMTP Port: 587" -ForegroundColor Yellow
Write-Host "   - Encryption: TLS/StartTLS" -ForegroundColor Yellow
Write-Host "   - Authentication: Required" -ForegroundColor Yellow
Write-Host ""
Write-Host "2. For local applications that need to send email, use the Send-MailMessage cmdlet:" -ForegroundColor Yellow
Write-Host '   $securePassword = ConvertTo-SecureString "YourPassword" -AsPlainText -Force' -ForegroundColor Yellow
Write-Host '   $credential = New-Object System.Management.Automation.PSCredential("your-account@yourdomain.com", $securePassword)' -ForegroundColor Yellow
Write-Host '   Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -Credential $credential -From "sender@yourdomain.com" -To "recipient@example.com" -Subject "Test" -Body "Test message"' -ForegroundColor Yellow