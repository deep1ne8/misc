# Script to set a specific account as default in Outlook
# This uses the Outlook COM object to programmatically change the default account

# Define the email address to set as default
$targetEmailAddress = "theexecutivedirector@far-roundtable.org"

try {
    # Create Outlook COM object
    $outlook = New-Object -ComObject Outlook.Application
    Write-Host "Successfully connected to Outlook"
    
    # Get Outlook accounts
    $accounts = $outlook.Session.Accounts
    Write-Host "Found $($accounts.Count) accounts"
    
    # Find the target account
    $targetAccount = $null
    foreach ($account in $accounts) {
        Write-Host "Checking account: $($account.DisplayName) - $($account.SmtpAddress)"
        if ($account.SmtpAddress -eq $targetEmailAddress) {
            $targetAccount = $account
            Write-Host "Found target account: $($account.DisplayName)"
            break
        }
    }
    
    if ($null -eq $targetAccount) {
        Write-Host "Error: Account with email address '$targetEmailAddress' not found!" -ForegroundColor Red
        exit 1
    }
    
    # Set the account as default
    # This uses the DeliveryStore property to set the default account
    $outlook.Session.DefaultStore = $targetAccount.DeliveryStore
    Write-Host "Successfully set '$targetEmailAddress' as default account" -ForegroundColor Green
    
    # Display a confirmation message box
    $shell = New-Object -ComObject WScript.Shell
    $shell.Popup("Default email account has been set to: $targetEmailAddress", 0, "Default Account Changed", 64)
    
    # Clean up COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    # Display error message box
    $shell = New-Object -ComObject WScript.Shell
    $shell.Popup("Error: $($_.Exception.Message)", 0, "Error Setting Default Account", 16)
}