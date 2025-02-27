<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 24.09.17.1810

#
# .DESCRIPTION
# This script runs a variety of cmdlets to establish a baseline of the sharing status of a Calendar.
#
# .PARAMETER Identity
#  Owner Mailbox to query, owner of the Mailbox sharing the calendar.
#  Receiver of the shared mailbox, often the Delegate.
#
# .EXAMPLE
# Check-SharingStatus.ps1 -Owner Owner@contoso.com -Receiver Receiver@contoso.com

# Define the parameters
[CmdletBinding()] 
param(
    [Parameter(Mandatory=$true)]
    [string]$Owner = "jeff@distributedsun.com",
    [Parameter(Mandatory=$true)]
    [string]$Receiver = "kate@trucurrent.com"
)

$BuildVersion = "24.09.17.1810"





function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "24.09.17.1810"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path ((Get-Item $env:TEMP).FullName) $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}

if (Test-ScriptVersion -AutoUpdate) {
    # Update was downloaded, so stop here.
    Write-Host "Script was updated. Please rerun the command."  -ForegroundColor Yellow
    return
}

Write-Verbose "Script Versions: $BuildVersion"

$script:PIIAccess = $true #Assume we have PII access until we find out otherwise

<#
.SYNOPSIS
    Formats the CalendarSharingInvite logs from Export-MailboxDiagnosticLogs for a given identity.
.DESCRIPTION
    This function processes calendar sharing accept logs for a given identity and outputs the most recent update for each recipient.
.PARAMETER Identity
    The SMTP Address for which to process calendar sharing accept logs.
#>
function ProcessCalendarSharingInviteLogs {
    param (
        [string]$Identity
    )

    # Define the header row
    $header = "Timestamp", "Mailbox", "Entry MailboxOwner", "Recipient", "RecipientType", "SharingType", "DetailLevel"
    $csvString = @()
    $csvString = $header -join ","
    $csvString += "`n"

    # Call the Export-MailboxDiagnosticLogs cmdlet and store the output in a variable
    try {
        # Call the Export-MailboxDiagnosticLogs cmdlet and store the output in a variable
        # -ErrorAction is not supported on Export-MailboxDiagnosticLogs
        # $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName CalendarSharingInvite -ErrorAction SilentlyContinue

        $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName CalendarSharingInvite
    } catch {
        # Code to run if an error occurs
        Write-Error "An error occurred: $_"
    }

    # check if the output is empty
    if ($null -eq $logOutput.MailboxLog) {
        Write-Host "No data found for [$Identity]."
        return
    }

    $logLines =@()
    # Split the output into an array of lines
    $logLines = $logOutput.MailboxLog -split "`r`n"

    # Loop through each line of the output
    foreach ($line in $logLines) {
        if ($line -like "*RecipientType*") {
            $csvString += $line + "`n"
        }
    }

    # Clean up output
    $csvString = $csvString.Replace("Mailbox: ", "")
    $csvString = $csvString.Replace("Entry MailboxOwner:", "")
    $csvString = $csvString.Replace("Recipient:", "")
    $csvString = $csvString.Replace("RecipientType:", "")
    $csvString = $csvString.Replace("Handler=", "")
    $csvString = $csvString.Replace("ms-exchange-", "")
    $csvString = $csvString.Replace("DetailLevel=", "")

    # Convert the CSV string to an object
    $csvObject = $csvString | ConvertFrom-Csv

    # Access the values as properties of the object
    foreach ($row in $csvObject) {
        Write-Debug "$($row.Recipient) - $($row.SharingType) - $($row.detailLevel)"
    }

    #Filter the output to get the most recent update foreach recipient
    $mostRecentRecipients = $csvObject | Sort-Object Recipient -Unique | Sort-Object Timestamp -Descending

    # Output the results to the console
    Write-Host "User [$Identity] has shared their calendar with the following recipients:"
    $mostRecentRecipients | Format-Table -a Timestamp, Recipient, SharingType, DetailLevel
}

<#
.SYNOPSIS
    Formats the AcceptCalendarSharingInvite logs from Export-MailboxDiagnosticLogs for a given identity.
.DESCRIPTION
    This function processes calendar sharing invite logs.
.PARAMETER Identity
    The SMTP Address for which to process calendar sharing accept logs.
#>
function ProcessCalendarSharingAcceptLogs {
    param (
        [string]$Identity
    )

    # Define the header row
    $header = "Timestamp", "Mailbox", "SharedCalendarOwner", "FolderName"
    $csvString = @()
    $csvString = $header -join ","
    $csvString += "`n"

    # Call the Export-MailboxDiagnosticLogs cmdlet and store the output in a variable
    try {
        # Call the Export-MailboxDiagnosticLogs cmdlet and store the output in a variable
        # -ErrorAction is not supported on Export-MailboxDiagnosticLogs
        # $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName AcceptCalendarSharingInvite -ErrorAction SilentlyContinue

        $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName AcceptCalendarSharingInvite
    } catch {
        # Code to run if an error occurs
        Write-Error "An error occurred: $_"
    }

    # check if the output is empty
    if ($null -eq $logOutput.MailboxLog) {
        Write-Host "No AcceptCalendarSharingInvite Logs found for [$Identity]."
        return
    }

    $logLines =@()
    # Split the output into an array of lines
    $logLines = $logOutput.MailboxLog -split "`r`n"

    # Loop through each line of the output
    foreach ($line in $logLines) {
        if ($line -like "*CreateInternalSharedCalendarGroupEntry*") {
            $csvString += $line + "`n"
        }
    }

    # Clean up output
    $csvString = $csvString.Replace("Mailbox: ", "")
    $csvString = $csvString.Replace("Entry MailboxOwner:", "")
    $csvString = $csvString.Replace("Entry CreateInternalSharedCalendarGroupEntry: ", "")
    $csvString = $csvString.Replace("Creating a shared calendar for ", "")
    $csvString = $csvString.Replace("calendar name ", "")

    # Convert the CSV string to an object
    $csvObject = $csvString | ConvertFrom-Csv

    # Access the values as properties of the object
    foreach ($row in $csvObject) {
        Write-Debug "$($row.Timestamp) - $($row.SharedCalendarOwner) - $($row.FolderName) "
    }

    # Filter the output to get the most recent update for each recipient
    # $mostRecentSharedCalendars = $csvObject |sort-object SharedCalendarOwner -Unique | Sort-Object Timestamp -Descending

    # Output the results to the console
    Write-Host "Receiver [$Identity] has accepted copies of the shared calendar from the following recipients on these dates:"
    #Write-Host $csvObject | Format-Table -a Timestamp, SharedCalendarOwner, FolderName
    $csvObject | Format-Table -a Timestamp, SharedCalendarOwner, FolderName
}

<#
.SYNOPSIS
    Formats the InternetCalendar logs from Export-MailboxDiagnosticLogs for a given identity.
.DESCRIPTION
    This function processes calendar sharing invite logs.
.PARAMETER Identity
    The SMTP Address for which to process calendar sharing accept logs.
#>
function ProcessInternetCalendarLogs {
    param (
        [string]$Identity
    )

    # Define the header row
    $header = "Timestamp", "Mailbox", "SyncDetails", "PublishingUrl", "RemoteFolderName", "LocalFolderId", "Folder"

    $csvString = @()
    $csvString = $header -join ","
    $csvString += "`n"

    try {
        # Call the Export-MailboxDiagnosticLogs cmdlet and store the output in a variable
        # -ErrorAction is not supported on Export-MailboxDiagnosticLogs
        # $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName AcceptCalendarSharingInvite -ErrorAction SilentlyContinue

        $logOutput = Export-MailboxDiagnosticLogs $Identity -ComponentName InternetCalendar
    } catch {
        # Code to run if an error occurs
        Write-Error "An error occurred: $_"
    }

    # check if the output is empty
    if ($null -eq $logOutput.MailboxLog) {
        Write-Host "No InternetCalendar Logs found for [$Identity]."
        Write-Host -ForegroundColor Yellow "User [$Identity] is not receiving any Published Calendars."
        return
    }

    $logLines =@()

    # Split the output into an array of lines
    $logLines = $logOutput.MailboxLog -split "`r`n"

    # Loop through each line of the output
    foreach ($line in $logLines) {
        if ($line -like "*Entry Sync Details for InternetCalendar subscription DataType=calendar*") {
            $csvString += $line + "`n"
        }
    }

    # Clean up output
    $csvString = $csvString.Replace("Mailbox: ", "")
    $csvString = $csvString.Replace("Entry Sync Details for InternetCalendar subscription DataType=calendar", "InternetCalendar")
    $csvString = $csvString.Replace("PublishingUrl=", "")
    $csvString = $csvString.Replace("RemoteFolderName=", "")
    $csvString = $csvString.Replace("LocalFolderId=", "")
    $csvString = $csvString.Replace("folder ", "")

    # Convert the CSV string to an object
    $csvObject = $csvString | ConvertFrom-Csv

    # Clean up the Folder column
    foreach ($row in $csvObject) {
        $row.Folder = $row.Folder.Split("with")[0]
    }

    Write-Host -ForegroundColor Cyan "Receiver [$Identity] is/was receiving the following Published Calendars:"
    $csvObject | Sort-Object -Unique RemoteFolderName | Format-Table -a RemoteFolderName, Folder, PublishingUrl
}

<#
.SYNOPSIS
    Display Calendar Owner information.
.DESCRIPTION
    This function displays key Calendar Owner information.
.PARAMETER Identity
    The SMTP Address for Owner of the shared calendar.
#>
function GetOwnerInformation {
    param (
        [string]$Owner
    )
    #Standard Owner information
    Write-Host -ForegroundColor DarkYellow "------------------------------------------------"
    Write-Host -ForegroundColor DarkYellow "Key Owner Mailbox Information:"
    Write-Host -ForegroundColor DarkYellow "`t Running 'Get-Mailbox $Owner'"
    $script:OwnerMB = Get-Mailbox $Owner
    # Write-Host "`t DisplayName:" $script:OwnerMB.DisplayName
    # Write-Host "`t Database:" $script:OwnerMB.Database
    # Write-Host "`t ServerName:" $script:OwnerMB.ServerName
    # Write-Host "`t LitigationHoldEnabled:" $script:OwnerMB.LitigationHoldEnabled
    # Write-Host "`t CalendarVersionStoreDisabled:" $script:OwnerMB.CalendarVersionStoreDisabled
    # Write-Host "`t CalendarRepairDisabled:" $script:OwnerMB.CalendarRepairDisabled
    # Write-Host "`t RecipientTypeDetails:" $script:OwnerMB.RecipientTypeDetails
    # Write-Host "`t RecipientType:" $script:OwnerMB.RecipientType

    if (-not $script:OwnerMB) {
        Write-Host -ForegroundColor Yellow "Could not find Owner Mailbox [$Owner]."
        Write-Host -ForegroundColor DarkYellow "Defaulting to External Sharing or Publishing."
        return
    }

    $script:OwnerMB | Format-List DisplayName, Database, ServerName, LitigationHoldEnabled, CalendarVersionStoreDisabled, CalendarRepairDisabled, RecipientType*

    if ($null -eq $script:OwnerMB) {
        Write-Host -ForegroundColor Red "Could not find Owner Mailbox [$Owner]."
        exit
    }

    Write-Host -ForegroundColor DarkYellow "Send on Behalf Granted to :"
    foreach ($del in $($script:OwnerMB.GrantSendOnBehalfTo)) {
        Write-Host -ForegroundColor Blue "`t$($del)"
    }
    Write-Host "`n`n`n"

    if ($script:OwnerMB.DisplayName -like "Redacted*") {
        Write-Host -ForegroundColor Yellow "Do Not have PII information for the Owner."
        Write-Host -ForegroundColor Yellow "Get PII Access for $($script:OwnerMB.Database)."
        $script:PIIAccess = $false
    }

    Write-Host -ForegroundColor DarkYellow "Owner Calendar Folder Statistics:"
    Write-Host -ForegroundColor DarkYellow "`t Running 'Get-MailboxFolderStatistics -Identity $Owner -FolderScope Calendar'"
    $OwnerCalendar = Get-MailboxFolderStatistics -Identity $Owner -FolderScope Calendar
    $OwnerCalendarName = ($OwnerCalendar | Where-Object FolderType -EQ "Calendar").Name

    $OwnerCalendar | Format-Table -a FolderPath, ItemsInFolder, FolderAndSubfolderSize

    Write-Host -ForegroundColor DarkYellow "Owner Calendar Permissions:"
    Write-Host -ForegroundColor DarkYellow "`t Running 'Get-mailboxFolderPermission "${Owner}:\$OwnerCalendarName" | Format-Table -a User, AccessRights, SharingPermissionFlags'"
    Get-mailboxFolderPermission "${Owner}:\$OwnerCalendarName" | Format-Table -a User, AccessRights, SharingPermissionFlags

    Write-Host -ForegroundColor DarkYellow "Owner Root MB Permissions:"
    Write-Host -ForegroundColor DarkYellow "`t Running 'Get-mailboxPermission $Owner | Format-Table -a User, AccessRights, SharingPermissionFlags'"
    Get-mailboxPermission $Owner | Format-Table -a User, AccessRights, SharingPermissionFlags

    Write-Host -ForegroundColor DarkYellow "Owner Modern Sharing Sent Invites"
    ProcessCalendarSharingInviteLogs -Identity $Owner

    Write-Host -ForegroundColor DarkYellow "Owner Calendar Folder Information:"
    Write-Host -ForegroundColor DarkYellow "`t Running 'Get-MailboxCalendarFolder "${Owner}:\$OwnerCalendarName"'"

    $OwnerCalendarFolder = Get-MailboxCalendarFolder "${Owner}:\$OwnerCalendarName"
    if ($OwnerCalendarFolder.PublishEnabled) {
        Write-Host -ForegroundColor Green "Owner Calendar is Published."
        $script:OwnerPublished = $true
    } else {
        Write-Host -ForegroundColor Yellow "Owner Calendar is not Published."
        $script:OwnerPublished = $false
    }

    if ($OwnerCalendarFolder.ExtendedFolderFlags.Contains("SharedOut")) {
        Write-Host -ForegroundColor Green "Owner Calendar is Shared Out using Modern Sharing."
        $script:OwnerModernSharing = $true
    } else {
        Write-Host -ForegroundColor Yellow "Owner Calendar is not Shared Out."
        $script:OwnerModernSharing = $false
    }
}

<#
.SYNOPSIS
    Displays key information from the receiver of the shared Calendar.
.DESCRIPTION
    This function displays key Calendar Receiver information.
.PARAMETER Identity
    The SMTP Address for Receiver of the shared calendar.
#>
function GetReceiverInformation {
    param (
        [string]$Receiver
    )
    #Standard Receiver information
    Write-Host -ForegroundColor Cyan "`r`r`r------------------------------------------------"
    Write-Host -ForegroundColor Cyan "Key Receiver MB Information: [$Receiver]"
    Write-Host -ForegroundColor Cyan "Running: 'Get-Mailbox $Receiver'"
    $script:ReceiverMB = Get-Mailbox $Receiver

    if (-not $script:ReceiverMB) {
        Write-Host -ForegroundColor Yellow "Could not find Receiver Mailbox [$Receiver]."
        Write-Host -ForegroundColor Yellow "Defaulting to External Sharing or Publishing."
        return
    }

    $script:ReceiverMB | Format-List DisplayName, Database, LitigationHoldEnabled, CalendarVersionStoreDisabled, CalendarRepairDisabled, RecipientType*

    if ($script:OwnerMB.OrganizationalUnitRoot -eq $script:ReceiverMB.OrganizationalUnitRoot) {
        Write-Host -ForegroundColor Yellow "Owner and Receiver are in the same OU."
        Write-Host -ForegroundColor Yellow "Owner and Receiver will be using Internal Sharing."
        $script:SharingType = "InternalSharing"
    } else {
        Write-Host -ForegroundColor Yellow "Owner and Receiver are in different OUs."
        Write-Host -ForegroundColor Yellow "Owner and Receiver will be using External Sharing or Publishing."
        $script:SharingType = "ExternalSharing"
    }

    Write-Host -ForegroundColor Cyan "Receiver Calendar Folders (look for a copy of [$($OwnerMB.DisplayName)] Calendar):"
    Write-Host -ForegroundColor Cyan "Running: 'Get-MailboxFolderStatistics -Identity $Receiver -FolderScope Calendar'"
    $CalStats = Get-MailboxFolderStatistics -Identity $Receiver -FolderScope Calendar
    $CalStats | Format-Table -a FolderPath, ItemsInFolder, FolderAndSubfolderSize
    $ReceiverCalendarName = ($CalStats | Where-Object FolderType -EQ "Calendar").Name

    # Note $Owner has a * at the end in case we have had multiple setup for the same user, they will be appended with a " 1", etc.
    if (($CalStats | Where-Object Name -Like $owner*) -or ($CalStats | Where-Object Name -Like "$($ownerMB.DisplayName)*" )) {
        Write-Host -ForegroundColor Green "Looks like we might have found a copy of the Owner Calendar in the Receiver Mailbox."
        Write-Host -ForegroundColor Green "This is a good indication the there is a Modern Sharing Relationship between these users."
        Write-Host -ForegroundColor Green "If the clients use the Modern Sharing or not is a up to the client."
        $script:ModernSharing = $true

        $CalStats | Where-Object Name -Like $owner* | Format-Table -a FolderPath, ItemsInFolder, FolderAndSubfolderSize
        if (($CalStats | Where-Object Name -Like $owner*).count -gt 1) {
            Write-Host -ForegroundColor Yellow "Warning: Might have found more than one copy of the Owner Calendar in the Receiver Mailbox."
        }
    } else {
        Write-Host -ForegroundColor Yellow "Warning: Could not Identify the Owner's [$Owner] Calendar in the Receiver Mailbox."
    }

    if ($ReceiverCalendarName -like "REDACTED-*" ) {
        Write-Host -ForegroundColor Yellow "Do Not have PII information for the Receiver"
        $script:PIIAccess = $false
    }

    ProcessCalendarSharingAcceptLogs -Identity $Receiver
    ProcessInternetCalendarLogs -Identity $Receiver

    if (($script:SharingType -like "InternalSharing") -or
    ($script:SharingType -like "ExternalSharing")) {
        # Validate Modern Sharing Status
        if (Get-Command -Name Get-CalendarEntries -ErrorAction SilentlyContinue) {
            Write-Verbose "Found Get-CalendarEntries cmdlet. Running cmdlet: Get-CalendarEntries -Identity $Receiver"
            # ToDo: Check each value for proper sharing permissions (i.e.  $X.CalendarSharingPermissionLevel -eq "ReadWrite" )
            $ReceiverCalEntries = Get-CalendarEntries -Identity $Receiver
            # Write-Host "CalendarGroupName : $($ReceiverCalEntries.CalendarGroupName)"
            # Write-Host "CalendarName : $($ReceiverCalEntries.CalendarName)"
            # Write-Host "OwnerEmailAddress : $($ReceiverCalEntries.OwnerEmailAddress)"
            # Write-Host "SharingModelType: $($ReceiverCalEntries.SharingModelType)"
            # Write-Host "IsOrphanedEntry: $($ReceiverCalEntries.IsOrphanedEntry)"

            Write-Host -ForegroundColor Cyan "`r`r`r------------------------------------------------"
            Write-Host "New Model Calendar Sharing Entries:"
            $ReceiverCalEntries | Where-Object SharingModelType -Like New | Format-Table CalendarGroupName, CalendarName, OwnerEmailAddress, SharingModelType, IsOrphanedEntry

            Write-Host -ForegroundColor Cyan "`r`r`r------------------------------------------------"
            Write-Host "Old Model Calendar Sharing Entries:"
            Write-Host "Consider upgrading these to the new model."
            $ReceiverCalEntries | Where-Object SharingModelType -Like Old | Format-Table CalendarGroupName, CalendarName, OwnerEmailAddress, SharingModelType, IsOrphanedEntry

            # need to check if Get-CalendarValidationResult in the PS Workspace
            if ((Get-Command -Name Get-CalendarValidationResult -ErrorAction SilentlyContinue) -and
                $null -ne $ReceiverCalEntries) {
                $ewsId_del= $ReceiverCalEntries[0].LocalFolderId
                Write-Host "Running cmdlet: Get-CalendarValidationResult -Version V2 -Identity $Receiver -SourceCalendarId $ewsId_del -TargetUserId $Owner -IncludeAnalysis 1 -OnlyReportErrors 1 | FT -a GlobalObjectId, EventValidationResult  "
                Get-CalendarValidationResult -Version V2 -Identity $Receiver -SourceCalendarId $ewsId_del -TargetUserId $Owner -IncludeAnalysis 1 -OnlyReportErrors 1 | Format-List UserPrimarySMTPAddress, Subject, GlobalObjectId, EventValidationResult, EventComparisonResult
            }
        }

        #Output key Modern Sharing information
        if (($script:PIIAccess) -and (-not ([string]::IsNullOrEmpty($script:OwnerMB)))) {
            Write-Host "Checking for Owner copy Calendar in Receiver Calendar:"
            Write-Host "Running cmdlet:"
            Write-Host -NoNewline -ForegroundColor Yellow "Get-MailboxCalendarFolder -Identity ${Receiver}:\$ReceiverCalendarName\$($script:OwnerMB.DisplayName)"
            try {
                Get-MailboxCalendarFolder -Identity "${Receiver}:\$ReceiverCalendarName\$($script:OwnerMB.DisplayName)" | Format-List Identity, CreationTime, ExtendedFolderFlags, ExtendedFolderFlags2, CalendarSharingFolderFlags, CalendarSharingOwnerSmtpAddress, CalendarSharingPermissionLevel, SharingLevelOfDetails, SharingPermissionFlags, LastAttemptedSyncTime, LastSuccessfulSyncTime, SharedCalendarSyncStartDate
            } catch {
                Write-Error "Failed to get the Owner Calendar from the Receiver Mailbox.  This is fine if not using Modern Sharing."
            }
        } else {
            Write-Host "Do Not have PII information for the Owner, so can not check the Receivers Copy of the Owner Calendar."
            Write-Host "Get PII Access for both mailboxes and try again."
        }
    }
}

# Main
$script:ModernSharing
$script:SharingType
GetOwnerInformation -Owner $Owner
GetReceiverInformation -Receiver $Receiver

Write-Host -ForegroundColor Blue "`r`r`r------------------------------------------------"
Write-Host -ForegroundColor Blue "Summary:"
Write-Host -ForegroundColor Blue "Mailbox Owner [$Owner] and Receiver [$Receiver] are using [$script:SharingType] for Calendar Sharing."
Write-Host -ForegroundColor Blue "It appears like the backend [$(if ($script:ModernSharing) {"IS"} else {"is NOT"})] using Modern Calendar Sharing."

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1bDlsiiFO2z0U
# 9m7/bW6vykzGVmSdLXMzi7gAy+stcaCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIG75EQ6EeRNdyA8vtIZDC9jG
# tPpeKDfGw47aM3Gl7qffMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEARIzE0pvjHoJk9bGhKOSawVSH4z/qR91LJn2/SIAU7VGZuUJ2Gr+Q8UpT
# a1blo94Ea9UuHcyv3o2FY+xjajbwgNZu1wRKLxx/OHKsS/HIlzfyKbRvT108Bbng
# UJ0Pc7cCFe7oJW7Kjt0ZC0Ax7R57e5nYv2kuH8IeS39theCAb/2JXK7fOd9XYFbx
# On/MFbbfP4Nvi6Q2XrWgW19fZssVXlOKpKrlq3NBpaun6nSwxI1HuuJpLcZurLUy
# aQRJiZjZbDjJWTmSxmL5MzRfUl3z7GI24WSxn3TIggmR6BtA0RpbwjcdHyrFKsND
# Eay6qZQLl7NfyukphMkJu+MlLIl/lKGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCBRJiM1eRovWTxvSyEu8Is+Ml4iRIOoSX12cDvQH5e2bgIGZ63adn69
# GBMyMDI1MDIxNDIyMTk0MC4yMjdaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTkzNS0w
# M0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAekPcTB+XfESNgABAAAB6TANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzEyMDYxODQ1
# MjZaFw0yNTAzMDUxODQ1MjZaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTkzNS0wM0UwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCsmowxQRVgp4TSc3nTa6yrAPJnV6A7aZYnTw/yx90u
# 1DSH89nvfQNzb+5fmBK8ppH76TmJzjHUcImd845A/pvZY5O8PCBu7Gq+x5Xe6plQ
# t4xwVUUcQITxklOZ1Rm9fJ5nh8gnxOxaezFMM41sDI7LMpKwIKQMwXDctYKvCyQy
# 6kO2sVLB62kF892ZwcYpiIVx3LT1LPdMt1IeS35KY5MxylRdTS7E1Jocl30NgcBi
# JfqnMce05eEipIsTO4DIn//TtP1Rx57VXfvCO8NSCh9dxsyvng0lUVY+urq/G8QR
# FoOl/7oOI0Rf8Qg+3hyYayHsI9wtvDHGnT30Nr41xzTpw2I6ZWaIhPwMu5DvdkEG
# zV7vYT3tb9tTviY3psul1T5D938/AfNLqanVCJtP4yz0VJBSGV+h66ZcaUJOxpbS
# IjImaOLF18NOjmf1nwDatsBouXWXFK7E5S0VLRyoTqDCxHG4mW3mpNQopM/U1WJn
# jssWQluK8eb+MDKlk9E/hOBYKs2KfeQ4HG7dOcK+wMOamGfwvkIe7dkylzm8BeAU
# QC8LxrAQykhSHy+FaQ93DAlfQYowYDtzGXqE6wOATeKFI30u9YlxDTzAuLDK073c
# ndMV4qaD3euXA6xUNCozg7rihiHUaM43Amb9EGuRl022+yPwclmykssk30a4Rp3v
# 9QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFJF+M4nFCHYjuIj0Wuv+jcjtB+xOMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBWsSp+rmsxFLe61AE90Ken2XPgQHJDiS4S
# bLhvzfVjDPDmOdRE75uQohYhFMdGwHKbVmLK0lHV1Apz/HciZooyeoAvkHQaHmLh
# wBGkoyAAVxcaaUnHNIUS9LveL00PwmcSDLgN0V/Fyk20QpHDEukwKR8kfaBEX83A
# yvQzlf/boDNoWKEgpdAsL8SzCzXFLnDozzCJGq0RzwQgeEBr8E4K2wQ2WXI/ZJxZ
# S/+d3FdwG4ErBFzzUiSbV2m3xsMP3cqCRFDtJ1C3/JnjXMChnm9bLDD1waJ7TPp5
# wYdv0Ol9+aN0t1BmOzCj8DmqKuUwzgCK9Tjtw5KUjaO6QjegHzndX/tZrY792dfR
# AXr5dGrKkpssIHq6rrWO4PlL3OS+4ciL/l8pm+oNJXWGXYJL5H6LNnKyXJVEw/1F
# bO4+Gz+U4fFFxs2S8UwvrBbYccVQ9O+Flj7xTAeITJsHptAvREqCc+/YxzhIKkA8
# 8Q8QhJKUDtazatJH7ZOdi0LCKwgqQO4H81KZGDSLktFvNRhh8ZBAenn1pW+5UBGY
# z2GpgcxVXKT1CuUYdlHR9D6NrVhGqdhGTg7Og/d/8oMlPG3YjuqFxidiIsoAw2+M
# hI1zXrIi56t6JkJ75J69F+lkh9myJJpNkx41sSB1XK2jJWgq7VlBuP1BuXjZ3qgy
# m9r1wv0MtTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE5MzUtMDNFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQCr
# aYf1xDk2rMnU/VJo2GGK1nxo8aCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA61mqVDAiGA8yMDI1MDIxNDExNDEw
# OFoYDzIwMjUwMjE1MTE0MTA4WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDrWapU
# AgEAMAcCAQACAhVYMAcCAQACAhMJMAoCBQDrWvvUAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAAh7lYMigDzAtVBYO2SljY+FmFI8yavOvMgRdQur66WHVj6D
# 0uilMmvdZGMZFogeHA+pGWplPEhMEhNCECV2ddXWjpEzl0jjsXFmlwExSlQHNd3X
# IJG2nOPKaq7uFjTqHJZHbf8oRM6P3vfyxmaSBvSyuTNSsAiEl6Wne/AdmRBtLpgY
# FKNhojXR0vV9mEmbcvBe2Xvhi0urKPIbcUwCD+VAOv9zUSdBFmCy1wWpys3ymhI7
# i/L8lOB4ChPcrjr5N6CA7/XFrJGxcutXuo9Q8R54gpcVtQ3kRaZuYfWxR29GUcL3
# kIpI9lJyv6JpyV4z5db/Mp4IYz3qDDaioJOIqXAxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAekPcTB+XfESNgABAAAB6TAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCAoVY+tLwBkgZdwr8ShJwDYCFGYWcpJELvcnTtDKWt/rDCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIKSQkniXaTcmj1TKQWF+x2U4riVo
# rGD8TwmgVbN9qsQlMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHpD3Ewfl3xEjYAAQAAAekwIgQgMFdc4e46bL7KefbHQ/QIV3gJq5DQ
# pqRJ6KTBeye4kFQwDQYJKoZIhvcNAQELBQAEggIAqEIfGC/N7oVIHz6CgChU63KD
# LR9cn/506uaHLlvIvdvH0FxCFOt3MZ0q3bOcCpvQdlJa9NESn2mZcgXTIt0w50q1
# QT15wQNUccusxBhOHdzuspmp79RGWrEprDE2osJxbrOctHxQrewa8+OMsXNm0czd
# bic2Ep0GjSYgs//41M3RsOubQyWyO6Tq6loQWi0W1mK1gzd/nBPU2F6My4eN/PGm
# FIyuhwmdx5Tn60Ez10hquXIAzJGzNEWD0i5iGqJAjxY4y0FfrMj4Ygvkeacz4gEe
# 2FF/AojhMMN7ldvwTRvL0xdSBexTsUPnbBFF2cYGTFgXswJ8JWFg05hZuW67VWQj
# zYg4pyq1V0hwlOM1QKyJsdALLG7/cYeU6E0o0z5cRf6cewti7SGaBdhpYhc+mXSx
# c9+jiqJpLf9Us/hZ1qVwdkDbtMLP/GhX7iwhsaXvUPD2V27H/l9lvCWaLxwEyKC5
# BGRmwgCacbSFhJhlfXbKKtOeTIbwdQjA5f23aj0FxOL5hQcnNgvqoHZg+5CuvHkW
# T5f3Z5Es7r9a2AJtppsDT11bgTaSLggloKB2p0oRjsQbmRYzHprjefSsmnLH/g3m
# NNsZqBSu2s5ebxdZyspJteLB0xSHiY/37qqDVIUl9+GqaL//MqbHK85NAXdcs+uA
# Jrlf8SjF0+E12GIYs1c=
# SIG # End signature block
