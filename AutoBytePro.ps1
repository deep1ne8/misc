#requires -Version 5.1
<#
.SYNOPSIS
    AutoByte Professional - PowerShell Task Automation App with LLM Chat Integration

.DESCRIPTION
    A comprehensive PowerShell application for IT automation that provides:
    - Script execution from a GitHub repository
    - Interactive terminal interface with visual enhancements
    - Integration with Llama LLM for AI-assisted troubleshooting
    - Robust error handling and logging capabilities

.NOTES
    Name: AutoByte-Pro.ps1
    Author: deep1ne8 (Enhanced by Claude)
    Version: 2.0
    Created: 2025-03-20
    
.EXAMPLE
    .\AutoByte-Pro.ps1
    Runs the main application with default settings

.EXAMPLE
    .\AutoByte-Pro.ps1 -Debug
    Runs the application with debug logging enabled
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$DebugMode,
    
    [string]$LlamaModelPath = "C:\Users\deep1ne\Mentor\models\llama-2-7b-chat.gguf",
    
    [string]$LogPath = "$env:TEMP\AutoByte\Logs"
)

#region Configuration
# Application settings
$script:AppName = "AutoByte Professional"
$script:Version = "2.0"
$script:LogFile = Join-Path -Path $LogPath -ChildPath "$(Get-Date -Format 'yyyyMMdd')-AutoByte.log"
$script:ScriptRepository = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts"
$script:DebugMode = $DebugMode.IsPresent

# Default console colors 
$script:Colors = @{
    'Primary' = [System.ConsoleColor]::Cyan
    'Secondary' = [System.ConsoleColor]::Blue
    'Success' = [System.ConsoleColor]::Green
    'Error' = [System.ConsoleColor]::Red
    'Warning' = [System.ConsoleColor]::Yellow
    'Info' = [System.ConsoleColor]::White
    'Accent' = [System.ConsoleColor]::Magenta
}

# Script definitions with metadata
$script:ScriptDefinitions = @(
    @{ 
        ScriptUrl = "$script:ScriptRepository/DiskCleaner.ps1"
        Description = "Disk Cleaner"
        Category = "System Maintenance"
        Details = "Removes temporary files and cleans up disk space"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/EnableFilesOnDemand.ps1"
        Description = "Enable Files On-Demand"
        Category = "Cloud Storage"
        Details = "Configures OneDrive Files On-Demand feature"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/DownloadandInstallPackage.ps1"
        Description = "Download & Install Package"
        Category = "Software Installation"
        Details = "Downloads and installs software packages from URLs"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/CheckUserProfileIssue.ps1"
        Description = "Check User Profile"
        Category = "Troubleshooting"
        Details = "Diagnoses and resolves common user profile issues"
        RequiresAdmin = $false
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/BloatWareRemover.ps1"
        Description = "Dell Bloatware Remover"
        Category = "System Cleanup"
        Details = "Removes pre-installed Dell bloatware applications"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/InstallWindowsUpdate.ps1"
        Description = "Reset & Install Windows Update"
        Category = "System Updates"
        Details = "Resets Windows Update components and installs pending updates"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/WindowsSystemRepair.ps1"
        Description = "Windows System Repair"
        Category = "System Repair"
        Details = "Runs system file checks and repairs Windows components"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/ResetandClearWindowsSearchDB.ps1"
        Description = "Reset Windows Search DB"
        Category = "System Repair"
        Details = "Resets and rebuilds the Windows Search database"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/InstallMSProjects.ps1"
        Description = "Install MS Projects"
        Category = "Software Installation"
        Details = "Installs Microsoft Project and related components"
        RequiresAdmin = $true
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/CheckDriveSpace.ps1"
        Description = "Check Drive Space"
        Category = "System Monitoring"
        Details = "Reports available space on all connected drives"
        RequiresAdmin = $false
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/InternetSpeedTest.ps1"
        Description = "Internet Speed Test"
        Category = "Network Diagnostics"
        Details = "Tests download and upload internet speeds"
        RequiresAdmin = $false
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/InternetLatencyTestNew.ps1"
        Description = "Internet Latency Test" 
        Category = "Network Diagnostics"
        Details = "Tests network latency to various endpoints"
        RequiresAdmin = $false
    },
    @{ 
        ScriptUrl = "$script:ScriptRepository/WorkPaperMonitorTroubleShooter.ps1"
        Description = "WorkPaper Monitor Troubleshooter"
        Category = "Application Support"
        Details = "Diagnoses and fixes issues with WorkPaper Monitor"
        RequiresAdmin = $false
    }
)
#endregion

#region Helper Functions
function Initialize-Environment {
    <#
    .SYNOPSIS
        Initializes the application environment
    #>
    [CmdletBinding()]
    param()
    
    try {
        # Create log directory if it doesn't exist
        try {
            if (-not (Test-Path -Path $LogPath)) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
                Write-Log "Created log directory: $LogPath" -Level INFO
            }
        } catch {
            Write-Host "Error creating log directory: $LogPath. Please check permissions." -ForegroundColor Red
            throw
        }
        
        # Initialize log file with header
        if (-not (Test-Path -Path $script:LogFile)) {
            $logHeader = @"
    #------------------------------------------------------------------------------
    # $script:AppName Log File
    # Version: $script:Version
    # Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    # Computer: $env:COMPUTERNAME
    # User: $env:USERNAME
    # PowerShell Version: $($PSVersionTable.PSVersion)
    #------------------------------------------------------------------------------
    "@
            $logHeader | Out-File -FilePath $script:LogFile -Encoding utf8 -Force
        }
#------------------------------------------------------------------------------
# $script:AppName Log File
# Version: $script:Version
# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Computer: $env:COMPUTERNAME
# User: $env:USERNAME
# PowerShell Version: $($PSVersionTable.PSVersion)
#------------------------------------------------------------------------------

"@
        }
        
        if (![string]::IsNullOrEmpty($LlamaModelPath)) {
            if (!(Test-Path $LlamaModelPath)) {
                Write-Log "Llama model not found at path: $LlamaModelPath" -Level WARNING
                Write-Host "Error: Llama model file not found. Please verify the path: $LlamaModelPath" -ForegroundColor Red
            }
        } else {
            Write-Log "Llama model path is empty. Please provide a valid path." -Level ERROR
            Write-Host "Error: Llama model path is not specified. Update the script parameters." -ForegroundColor Red
            throw "Llama model path is missing."
        }
        
        # Verify Llama model exists if specified
        if (![string]::IsNullOrEmpty($LlamaModelPath) -and !(Test-Path $LlamaModelPath)) {
            Write-Log "Llama model not found at path: $LlamaModelPath" -Level WARNING
        }
        
        # Check for internet connectivity
        $internetConnected = Test-InternetConnectivity
        if (!$internetConnected) {
            Write-Log "No internet connection detected. Some features may not work correctly." -Level WARNING
        }
        
        Write-Log "$script:AppName v$script:Version started" -Level INFO
    }
    catch {
        # Critical initialization failure
        $errorMessage = "Failed to initialize environment: $_"
        
        # Try to create minimal log even if normal logging fails
        try {
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            "[CRITICAL] [$timestamp] $errorMessage" | Out-File -FilePath "$env:TEMP\AutoByte_Error.log" -Append -Force
        }
        catch {
            # Last resort error handling
            Write-Host "CRITICAL ERROR: Could not initialize application environment." -ForegroundColor Red
            Write-Host $errorMessage -ForegroundColor Red
        }
        
        throw $errorMessage
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message to the log file and optionally to the console.
    
    .PARAMETER Message
        The message to be logged.
    
    .PARAMETER Level
        The severity level of the message (INFO, WARNING, ERROR, SUCCESS, DEBUG).
    
    .PARAMETER NoConsole
        If specified, the message will only be written to the log file and not to the console.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Position = 1)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO',
        
        [Parameter()]
        [switch]$NoConsole
    )
    
    # Skip DEBUG messages unless debug mode is enabled
    if ($Level -eq 'DEBUG' -and -not $script:DebugMode) {
        return
    }
    
    try {
        # Format timestamp for log entry
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Append to log file
        $logEntry | Out-File -FilePath $script:LogFile -Encoding utf8 -Append
        
        # Console color mapping
        $consoleColors = @{
            'INFO'    = $script:Colors.Info
            'WARNING' = $script:Colors.Warning
            'ERROR'   = $script:Colors.Error
            'SUCCESS' = $script:Colors.Success
            'DEBUG'   = $script:Colors.Secondary
        }
        
        # Write to console if not suppressed
        if (-not $NoConsole) {
            # Use Write-Host with configured console colors
            Write-Host $logEntry -ForegroundColor $consoleColors[$Level]
        }
    }
    catch {
        # Last resort error handling if logging itself fails
        try {
            Write-Host "[$Level] $Message" -ForegroundColor Red
            Write-Host "Failed to write to log file: $_" -ForegroundColor Red
        }
        catch {
            # If all else fails, at least try to output something
            [Console]::Error.WriteLine("[$Level] $Message")
            [Console]::Error.WriteLine("Logging error: $_")
        }
    }
}

function Get-ConsoleWidth {
    <#
    .SYNOPSIS
        Gets the console window width with fallback for different environments
    #>
    [CmdletBinding()]
    param()
    
    try {
        # First try the standard .NET way
        $width = [Console]::WindowWidth
        
        # Ensure minimum width
        if ($width -lt 80) {
            $width = 80
        }
        
        return $width
    }
    catch {
        # Fallback for environments where WindowWidth isn't available
        return 100  # Reasonable default
    }
}

function Clear-Console {
    <#
    .SYNOPSIS
        Clears the console screen in a platform-independent way
    #>
    [CmdletBinding()]
    param()
    
    try {
        Clear-Host
        
        # For Unix-like systems, ANSI escape sequence can be used
        if ($PSVersionTable.Platform -eq 'Unix') {
            [Console]::Write("`e[2J`e[H")
        }
    }
    catch {
        # Fallback - just output some new lines
        1..50 | ForEach-Object { Write-Host "" }
    }
}

function Show-Header {
    <#
    .SYNOPSIS
        Displays a standardized header with application title
    
    .PARAMETER Title
        Optional subtitle to display
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Title = ""
    )
    
    $consoleWidth = Get-ConsoleWidth
    $separator = "=" * $consoleWidth
    
    Clear-Console
    
    Write-Host "`n" -NoNewline
    Write-Host $separator -ForegroundColor $script:Colors.Primary
    
    $appHeader = "                        $script:AppName                           "
    Write-Host $appHeader -BackgroundColor $script:Colors.Secondary -ForegroundColor White
    
    if (![string]::IsNullOrEmpty($Title)) {
        $subHeader = "                        $Title                           "
        Write-Host $subHeader -BackgroundColor $script:Colors.Secondary -ForegroundColor White
    }
    
    Write-Host $separator -ForegroundColor $script:Colors.Primary
    Write-Host "`n" -NoNewline
    
    # ASCII art logo
    Write-Host "   ______           __           ____             __" -ForegroundColor $script:Colors.Primary
    Write-Host "  /\  _  \         /\ \__       /\  _ \          /\ \__" -ForegroundColor White
    Write-Host "  \ \ \L\ \  __  __\ \  _\   __ \ \ \_\ \      __\ \  _\   __" -ForegroundColor $script:Colors.Error
    Write-Host "   \ \  __ \/\ \/\ \\ \ \/  / __ \ \  _ < /\ \/\ \\ \ \/  / __ \" -ForegroundColor $script:Colors.Error
    Write-Host "    \ \ \/\ \ \ \_\ \\ \ \_/\ \L\ \ \ \L\ \ \ \_\ \\ \ \_/\  __/" -ForegroundColor White
    Write-Host "     \ \_\ \_\ \____/ \ \__\ \____/\ \____/\/ ____ \\ \__\ \____\" -ForegroundColor $script:Colors.Primary
    Write-Host "      \/_/\/_/\/___/   \/__/\/___/  \/___/   /___/> \\/__/\/____/" -ForegroundColor $script:Colors.Primary
    Write-Host "                                               /\___/" -ForegroundColor White
    Write-Host "                                               \/__/" -ForegroundColor $script:Colors.Error
    Write-Host "`n" -NoNewline
    
    # Display application info
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $adminStatus = if ($isAdmin) { "Administrator" } else { "Standard User" }
    
    Write-Host "System: " -NoNewline -ForegroundColor White
    Write-Host "$env:COMPUTERNAME" -ForegroundColor $script:Colors.Primary
    
    Write-Host "User: " -NoNewline -ForegroundColor White
    Write-Host "$env:USERNAME ($adminStatus)" -ForegroundColor $script:Colors.Primary
    
    Write-Host "Version: " -NoNewline -ForegroundColor White
    Write-Host "v$script:Version" -ForegroundColor $script:Colors.Primary
    
    Write-Host "`n" -NoNewline
    
    # Additional subtitle if provided
    if (![string]::IsNullOrEmpty($Title)) {
        Write-Host $Title -BackgroundColor $script:Colors.Secondary -ForegroundColor White
        Write-Host "`n" -NoNewline
    }
    
    Write-Host $separator -ForegroundColor $script:Colors.Primary
}

function Show-Footer {
    <#
    .SYNOPSIS
        Displays standardized footer with navigation options
    
    .PARAMETER Options
        Array of options to display in the footer
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$Options = @("Press any key to continue...")
    )
    
    $consoleWidth = Get-ConsoleWidth
    $separator = "=" * $consoleWidth
    
    Write-Host "`n" -NoNewline
    Write-Host $separator -ForegroundColor $script:Colors.Primary
    
    foreach ($option in $Options) {
        Write-Host $option -ForegroundColor $script:Colors.Warning
    }
    
    Write-Host $separator -ForegroundColor $script:Colors.Primary
    Write-Host "`n" -NoNewline
}

function Test-InternetConnectivity {
    <#
    .SYNOPSIS
        Tests internet connectivity
    #>
    [CmdletBinding()]
    param()
    
    try {
        $testConnection = Test-Connection -ComputerName 'github.com' -Count 1 -Quiet
        return $testConnection
    }
    catch {
        # Fallback method if Test-Connection fails
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadString("https://github.com") | Out-Null
            return $true
        }
        catch {
            return $false
        }
    }
}

function Test-AdminPrivileges {
    <#
    .SYNOPSIS
        Checks if the current user has administrator privileges
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $isAdmin
}

function Show-Prompt {
    <#
    .SYNOPSIS
        Displays an input prompt with enhanced formatting
    
    .PARAMETER Message
        The prompt message to display
        
    .PARAMETER Options
        Array of valid options for input validation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter()]
        [string[]]$Options = @()
    )
    
    $optionsText = if ($Options.Count -gt 0) { " [" + ($Options -join "/") + "]" } else { "" }
    
    Write-Host "$Message$optionsText" -NoNewline -ForegroundColor $script:Colors.Warning
    $response = Read-Host
    
    # Validate against options if provided
    if ($Options.Count -gt 0 -and $response -notin $Options) {
        Write-Host "Invalid option. Please select from: $($Options -join ", ")" -ForegroundColor $script:Colors.Error
        return Show-Prompt -Message $Message -Options $Options
    }
    
    return $response
}

function Invoke-WithProgress {
    <#
    .SYNOPSIS
        Executes a scriptblock with a progress display
    
    .PARAMETER Action
        ScriptBlock to execute
        
    .PARAMETER ActionText
        Text to display during execution
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$Action,
        
        [Parameter()]
        [string]$ActionText = "Processing"
    )
    
    $oldProgress = $ProgressPreference
    $ProgressPreference = 'Continue'
    
    try {
        Write-Progress -Activity $ActionText -Status "Please wait..." -PercentComplete -1
        
        # Execute the action
        & $Action
        
        Write-Progress -Activity $ActionText -Completed
    }
    finally {
        $ProgressPreference = $oldProgress
    }
}
#endregion

#region Script Execution Functions
function Invoke-ScriptFromUrl {
    <#
    .SYNOPSIS
        Downloads and executes a script from a URL with error handling
    
    .PARAMETER Url
        The URL of the script to execute
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Url,
        
        [Parameter()]
        [string]$ScriptName = (Split-Path -Path $Url -Leaf)
    )
    
    try {
        Write-Log "Attempting to execute script from URL: $Url" -Level INFO
        
        # Check internet connectivity first
        if (!(Test-InternetConnectivity)) {
            throw "No internet connection available"
        }
        
        # Create a temporary file for the script
        $tempScriptPath = Join-Path -Path $env:TEMP -ChildPath "AutoByte_$([Guid]::NewGuid().ToString()).ps1"
        
        # Download the script with progress indication
        Write-Host "Downloading script..." -ForegroundColor $script:Colors.Info
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "AutoByte/$script:Version PowerShell/$($PSVersionTable.PSVersion)")
        
        try {
            $scriptContent = $webClient.DownloadString($Url)
            
            # Save to temporary file
            Set-Content -Path $tempScriptPath -Value $scriptContent -Force
            Write-Log "Script downloaded successfully to: $tempScriptPath" -Level DEBUG
        }
        catch {
            throw "Failed to download script: $_"
        }
        
        # Execute the script
        Write-Log "Executing script: $ScriptName" -Level INFO
        Write-Host "`n" -NoNewline
        Write-Host "Executing $ScriptName..." -ForegroundColor $script:Colors.Primary
        Write-Host "----------------------------------------------------------------" -ForegroundColor $script:Colors.Primary
        
        # Execute using dot sourcing to maintain scope
        & $tempScriptPath
        
        Write-Host "----------------------------------------------------------------" -ForegroundColor $script:Colors.Primary
        Write-Log "Script execution completed successfully: $ScriptName" -Level SUCCESS
        
        # Clean up
        if (Test-Path -Path $tempScriptPath) {
            Remove-Item -Path $tempScriptPath -Force
        }
        
        return $true
    }
    catch {
        Write-Log "Error executing script from URL: $_" -Level ERROR
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor $script:Colors.Error
        
        # Clean up on failure
        if (Test-Path -Path $tempScriptPath) {
            Remove-Item -Path $tempScriptPath -Force -ErrorAction SilentlyContinue
        }
        
        return $false
    }
}

function Show-ScriptCategoryMenu {
    <#
    .SYNOPSIS
        Displays scripts grouped by category
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "Script Categories"
    
    # Get distinct categories
    $categories = $script:ScriptDefinitions | Select-Object -ExpandProperty Category -Unique | Sort-Object
    
    $index = 1
    $categoryMap = @{}
    
    Write-Host "Select a script category:" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    foreach ($category in $categories) {
        $categoryMap[$index.ToString()] = $category
        Write-Host "$index. $category" -ForegroundColor $script:Colors.Success
        $index++
    }
    
    Write-Host ""
    Write-Host "B. Back to Main Menu" -ForegroundColor $script:Colors.Warning
    Write-Host "X. Exit" -ForegroundColor $script:Colors.Error
    
    Show-Footer
    
    $categoryChoice = Show-Prompt -Message "Enter your choice" -Options (@($categoryMap.Keys) + @("B", "X"))
    
    switch ($categoryChoice.ToUpper()) {
        "B" { Show-MainMenu }
        "X" { Exit-Application }
        default {
            $selectedCategory = $categoryMap[$categoryChoice]
            Show-ScriptsInCategory -Category $selectedCategory
        }
    }
}

function Show-ScriptsInCategory {
    <#
    .SYNOPSIS
        Displays scripts in the selected category
    
    .PARAMETER Category
        The category to display scripts for
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Category
    )
    
    Show-Header -Title "Category: $Category"
    
    # Filter scripts by category
    $categoryScripts = $script:ScriptDefinitions | Where-Object { $_.Category -eq $Category }
    
    $index = 1
    $scriptMap = @{}
    
    Write-Host "Select a script to run:" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    # Check if we have admin rights
    $isAdmin = Test-AdminPrivileges
    
    foreach ($script in $categoryScripts) {
        $scriptMap[$index.ToString()] = $script
        
        # Format script entry differently based on admin requirements
        if ($script.RequiresAdmin -and -not $isAdmin) {
            Write-Host "$index. $($script.Description) (requires admin)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "$index. $($script.Description)" -ForegroundColor $script:Colors.Success
        }
        
        # Display script details
        Write-Host "   $($script.Details)" -ForegroundColor $script:Colors.Info
        $index++
    }
    
    Write-Host ""
    Write-Host "B. Back to Categories" -ForegroundColor $script:Colors.Warning
    Write-Host "M. Main Menu" -ForegroundColor $script:Colors.Warning
    Write-Host "X. Exit" -ForegroundColor $script:Colors.Error
    
    Show-Footer
    
    $scriptChoice = Show-Prompt -Message "Enter your choice" -Options (@($scriptMap.Keys) + @("B", "M", "X"))
    
    switch ($scriptChoice.ToUpper()) {
        "B" { Show-ScriptCategoryMenu }
        "M" { Show-MainMenu }
        "X" { Exit-Application }
        default {
            $selectedScript = $scriptMap[$scriptChoice]
            
            # Check admin requirements
            if ($selectedScript.RequiresAdmin -and -not $isAdmin) {
                Show-Header -Title "Administrator Required"
                Write-Host "The script [$($selectedScript.Description)] requires administrator privileges." -ForegroundColor $script:Colors.Error
                Write-Host "Please restart the application as an administrator to run this script." -ForegroundColor $script:Colors.Warning
                Write-Host ""
                Show-Footer -Options @("Press any key to return to the category menu...")
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                Show-ScriptsInCategory -Category $Category
            }
            else {
                # Execute the script
                $success = Invoke-ScriptFromUrl -Url $selectedScript.ScriptUrl -ScriptName $selectedScript.Description
                $success
                
                # Show return menu
                Show-ReturnMenu -PreviousCategory $Category
            }
        }
    }
}

function Show-AllScriptsMenu {
    <#
    .SYNOPSIS
        Displays all available scripts in a single menu
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "All Available Scripts"
    
    $index = 1
    $scriptMap = @{}
    
    Write-Host "Select a script to run:" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    # Check if we have admin rights
    $isAdmin = Test-AdminPrivileges
    
    foreach ($script in $script:ScriptDefinitions) {
        $scriptMap[$index.ToString()] = $script
        
        # Format script entry differently based on admin requirements
        if ($script.RequiresAdmin -and -not $isAdmin) {
            Write-Host "$index. $($script.Description) [requires admin]" -ForegroundColor DarkGray
        }
        else {
            Write-Host "$index. $($script.Description)" -ForegroundColor $script:Colors.Success
        }
        
        # Display optional category and details
        Write-Host "   Category: $($script.Category) | $($script.Details)" -ForegroundColor $script:Colors.Info
        $index++
    }
    
    Write-Host ""
    Write-Host "B. Back to Main Menu" -ForegroundColor $script:Colors.Warning
    Write-Host "X. Exit" -ForegroundColor $script:Colors.Error
    
    Show-Footer
    
    $scriptChoice = Show-Prompt -Message "Enter your choice" -Options (@($scriptMap.Keys) + @("B", "X"))
    
    switch ($scriptChoice.ToUpper()) {
        "B" { Show-MainMenu }
        "X" { Exit-Application }
        default {
            $selectedScript = $scriptMap[$scriptChoice]
            
            # Check admin requirements
            if ($selectedScript.RequiresAdmin -and -not $isAdmin) {
                Show-Header -Title "Administrator Required"
                Write-Host "The script [$($selectedScript.Description)] requires administrator privileges." -ForegroundColor $script:Colors.Error
                Write-Host "Please restart the application as an administrator to run this script." -ForegroundColor $script:Colors.Warning
                Write-Host ""
                Show-Footer -Options @("Press any key to return to the main menu...")
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                Show-AllScriptsMenu
            }
            else {
                # Execute the script
                $success = Invoke-ScriptFromUrl -Url $selectedScript.ScriptUrl -ScriptName $selectedScript.Description
                $success
                
                # Show return menu
                Show-ReturnMenu
            }
        }
    }
}

function Show-ReturnMenu {
    <#
    .SYNOPSIS
        Displays options after script execution
    
    .PARAMETER PreviousCategory
        Optional category to return to instead of main menu
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$PreviousCategory = ""
    )
    
    Write-Host ""
    Write-Host "Script execution completed." -ForegroundColor $script:Colors.Success
    
    if (![string]::IsNullOrEmpty($PreviousCategory)) {
        $returnOptions = @(
            "1. Return to $PreviousCategory scripts",
            "2. Return to category menu",
            "3. Return to main menu",
            "4. Exit"
        )
        
        Show-Footer -Options $returnOptions
        
        $returnChoice = Show-Prompt -Message "Enter your choice" -Options @("1", "2", "3")
        
        switch ($returnChoice) {
            "1" { Show-AllScriptsMenu }
            "2" { Show-MainMenu }
            "3" { Exit-Application }
            default { Show-MainMenu }
        }
    }
}
#endregion

#region LLM Integration Functions
function Invoke-LlamaLLM {
    <#
    .SYNOPSIS
        Invokes the Llama LLM for AI-assisted troubleshooting
    
    .PARAMETER Prompt
        The prompt to send to the LLM
        
    .PARAMETER Temperature
        Controls randomness in generation (0.0-1.0)
        
    .PARAMETER MaxTokens
        Maximum number of tokens to generate
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        
        [Parameter()]
        [ValidateRange(0.0, 1.0)]
        [double]$Temperature = 0.7,
        
        [Parameter()]
        [int]$MaxTokens = 500
    )
    
    try {
        # Check if Llama model exists
        if (![string]::IsNullOrEmpty($LlamaModelPath) -and !(Test-Path $LlamaModelPath)) {
            throw "Llama model not found at path: $LlamaModelPath"
        }
        
        # Format the prompt for Llama chat format
        $formattedPrompt = @"
<s>[INST] <<SYS>> 
You are an IT troubleshooting assistant integrated into AutoByte Professional. 
Provide concise, accurate solutions to technical problems.
Focus on practical steps that can be implemented immediately.
If you need more information to diagnose a problem, ask clarifying questions.
For PowerShell solutions, provide clean, efficient scripts that follow best practices.
<</SYS>>

$Prompt [/INST]
"@
        
        # Prepare the llama.cpp command with parameters
        # Removed unused variable $tempOutputFile
        $promptFile = Join-Path -Path $env:TEMP -ChildPath "llama_prompt_$([Guid]::NewGuid().ToString()).txt"
        
        # Save prompt to a file
        $formattedPrompt | Out-File -FilePath $promptFile -Encoding utf8
        
        # Build command based on OS
        if ($PSVersionTable.Platform -eq 'Unix' -or $PSVersionTable.OS -like "*Linux*" -or $PSVersionTable.OS -like "*Darwin*") {
            $llamaCommand = "llama -m `"$LlamaModelPath`" -f `"$promptFile`" -t $Temperature -n $MaxTokens"
        } else {
            # Assuming Windows and a path to llama.cpp executable
            $llamaExePath = "llama.exe"  # Adjust this path as needed
            $llamaCommand = "$llamaExePath -m `"$LlamaModelPath`" -f `"$promptFile`" -t $Temperature -n $MaxTokens"
        }
        
        Write-Log "Invoking Llama LLM with prompt: $Prompt" -Level DEBUG
        
        # Show a waiting message
        Write-Host "Processing your request with AI assistant..." -ForegroundColor $script:Colors.Info
        Write-Host "This may take a moment depending on your hardware." -ForegroundColor $script:Colors.Info
        
        # Execute the Llama command
        $result = Invoke-Expression $llamaCommand
        
        # Clean up the prompt file
        if (Test-Path $promptFile) {
            Remove-Item $promptFile -Force -ErrorAction SilentlyContinue
        }
        
        # Extract the actual response (after the prompt)
        $response = $result -join "`n"
        
        # Basic parsing to extract just the response part
        $responseStartIndex = $response.IndexOf("[/INST]")
        if ($responseStartIndex -gt 0) {
            $responseText = $response.Substring($responseStartIndex + 7).Trim()
        } else {
            $responseText = $response.Trim()
        }
        
        return $responseText
    }
    catch {
        Write-Log "Error invoking Llama LLM: $_" -Level ERROR
        
        return @"
I apologize, but I encountered an error while processing your request.

Error details: $($_.Exception.Message)

You may want to:
1. Check if the Llama model file exists at the specified path
2. Ensure llama.cpp is properly installed and accessible
3. Try again with a simpler query
"@
    }
}

function Show-AIAssistantChat {
    <#
    .SYNOPSIS
        Opens an interactive chat session with the AI assistant
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "AI Troubleshooting Assistant"
    
    Write-Host "Welcome to the AI Troubleshooting Assistant powered by Llama LLM!" -ForegroundColor $script:Colors.Success
    Write-Host "Ask technical questions or describe issues you're experiencing." -ForegroundColor $script:Colors.Info
    Write-Host "Type 'exit', 'quit', or 'back' to return to the main menu." -ForegroundColor $script:Colors.Warning
    Write-Host ""
    
    $chatHistory = @()
    $exitChat = $false
    
    while (-not $exitChat) {
        # Get user input
        Write-Host "You: " -NoNewline -ForegroundColor $script:Colors.Primary
        $userQuery = Read-Host
        
        # Check for exit commands
        if ($userQuery -in @('exit', 'quit', 'back')) {
            $exitChat = $true
            continue
        }
        
        # Append chat history context if we have previous exchanges
        $fullPrompt = $userQuery
        if ($chatHistory.Count -gt 0) {
            $contextPrompt = "Previous conversation:`n"
            foreach ($exchange in $chatHistory) {
                $contextPrompt += "User: $($exchange.User)`nAssistant: $($exchange.Assistant)`n`n"
            }
            $contextPrompt += "Now answer this new question: $userQuery"
            $fullPrompt = $contextPrompt
        }
        
        # Get response from LLM
        $response = Invoke-LlamaLLM -Prompt $fullPrompt
        
        # Display response with formatting
        Write-Host "`nAssistant:" -ForegroundColor $script:Colors.Secondary
        
        # Format code blocks in the response
        $inCodeBlock = $false
        foreach ($line in $response -split "`n") {
            if ($line -match '```(powershell|ps|bash|cmd|batch)?') {
                $inCodeBlock = $true
                Write-Host ""  # Add spacing before code block
                continue
            }
            elseif ($line -match '```' -and $inCodeBlock) {
                $inCodeBlock = $false
                Write-Host ""  # Add spacing after code block
                continue
            }
            
            if ($inCodeBlock) {
                Write-Host $line -ForegroundColor $script:Colors.Accent
            }
            else {
                Write-Host $line -ForegroundColor $script:Colors.Info
            }
        }
        
        Write-Host ""  # Add spacing after response
        
        # Add to chat history (limit to last 5 exchanges to avoid context overflow)
        $chatHistory += @{ User = $userQuery; Assistant = $response }
        if ($chatHistory.Count -gt 5) {
            $chatHistory = $chatHistory | Select-Object -Last 5
        }
    }
    
    Show-MainMenu
}
#endregion

#region Menu Functions
function Show-MainMenu {
    <#
    .SYNOPSIS
        Displays the main application menu
    #>
    [CmdletBinding()]
    param()
    
    Show-Header
    
    Write-Host "Main Menu" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    # Determine if we have admin rights
    $isAdmin = Test-AdminPrivileges
    if (-not $isAdmin) {
        Write-Host "NOTE: Some scripts require administrator privileges." -ForegroundColor $script:Colors.Warning
        Write-Host "      These scripts will be shown in gray and cannot be executed." -ForegroundColor $script:Colors.Warning
        Write-Host ""
    }
    
    # Menu options
    Write-Host "1. Browse Scripts by Category" -ForegroundColor $script:Colors.Success
    Write-Host "2. View All Scripts" -ForegroundColor $script:Colors.Success
    Write-Host "3. AI Troubleshooting Assistant" -ForegroundColor $script:Colors.Success
    Write-Host "4. System Information" -ForegroundColor $script:Colors.Success
    Write-Host "5. About AutoByte" -ForegroundColor $script:Colors.Success
    Write-Host "X. Exit" -ForegroundColor $script:Colors.Error
    
    Show-Footer
    
    # Get user choice
    $menuChoice = Show-Prompt -Message "Enter your choice" -Options @("1", "2", "3", "4", "5", "X")
    
    switch ($menuChoice.ToUpper()) {
        "1" { Show-ScriptCategoryMenu }
        "2" { Show-AllScriptsMenu }
        "3" { Show-AIAssistantChat }
        "4" { Show-SystemInformation }
        "5" { Show-AboutScreen }
        "X" { Exit-Application }
        default { Show-MainMenu }
    }
}

function Show-SystemInformation {
    <#
    .SYNOPSIS
        Displays detailed system information
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "System Information"
    
    Write-Host "Collecting system information..." -ForegroundColor $script:Colors.Info
    
    # Collect system information with progress
    $systemInfo = Invoke-WithProgress -ActionText "Collecting system information" -Action {
        try {
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $processor = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop
            $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
            
            $networkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapter | 
                Where-Object { $_.PhysicalAdapter -eq $true -and $_.NetEnabled -eq $true } | 
                Select-Object Name, MacAddress, Speed, AdapterType
            
            return @{
                OS = @{
                    Name = $osInfo.Caption
                    Version = $osInfo.Version
                    BuildNumber = $osInfo.BuildNumber
                    Architecture = $osInfo.OSArchitecture
                    LastBoot = $osInfo.LastBootUpTime
                    InstallDate = $osInfo.InstallDate
                }
                Computer = @{
                    Name = $computerSystem.Name
                    Manufacturer = $computerSystem.Manufacturer
                    Model = $computerSystem.Model
                    SystemType = $computerSystem.SystemType
                    TotalPhysicalMemory = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
                }
                Processor = @{
                    Name = $processor.Name
                    Cores = $processor.NumberOfCores
                    LogicalProcessors = $processor.NumberOfLogicalProcessors
                    MaxClockSpeed = $processor.MaxClockSpeed
                }
                Storage = $drives | ForEach-Object {
                    @{
                        Drive = $_.DeviceID
                        VolumeName = $_.VolumeName
                        Size = [math]::Round($_.Size / 1GB, 2)
                        FreeSpace = [math]::Round($_.FreeSpace / 1GB, 2)
                        PercentFree = [math]::Round(($_.FreeSpace / $_.Size) * 100, 2)
                    }
                }
                Network = $networkAdapters
            }
        }
        catch {
            Write-Log "Error collecting system information: $_" -Level ERROR
            return $null
        }
    }
    
    if ($null -eq $systemInfo) {
        Write-Host "Failed to collect complete system information." -ForegroundColor $script:Colors.Error
    }
    else {
        # OS Information
        Write-Host "OPERATING SYSTEM" -ForegroundColor $script:Colors.Primary
        Write-Host "  Name:           $($systemInfo.OS.Name)" -ForegroundColor $script:Colors.Info
        Write-Host "  Version:        $($systemInfo.OS.Version)" -ForegroundColor $script:Colors.Info
        Write-Host "  Build:          $($systemInfo.OS.BuildNumber)" -ForegroundColor $script:Colors.Info
        Write-Host "  Architecture:   $($systemInfo.OS.Architecture)" -ForegroundColor $script:Colors.Info
        Write-Host "  Last Boot:      $($systemInfo.OS.LastBoot)" -ForegroundColor $script:Colors.Info
        Write-Host "  Install Date:   $($systemInfo.OS.InstallDate)" -ForegroundColor $script:Colors.Info
        Write-Host ""
        
        # Computer Information
        Write-Host "COMPUTER" -ForegroundColor $script:Colors.Primary
        Write-Host "  Name:           $($systemInfo.Computer.Name)" -ForegroundColor $script:Colors.Info
        Write-Host "  Manufacturer:   $($systemInfo.Computer.Manufacturer)" -ForegroundColor $script:Colors.Info
        Write-Host "  Model:          $($systemInfo.Computer.Model)" -ForegroundColor $script:Colors.Info
        Write-Host "  System Type:    $($systemInfo.Computer.SystemType)" -ForegroundColor $script:Colors.Info
        Write-Host "  Memory:         $($systemInfo.Computer.TotalPhysicalMemory) GB" -ForegroundColor $script:Colors.Info
        Write-Host ""
        
        # Processor Information
        Write-Host "PROCESSOR" -ForegroundColor $script:Colors.Primary
        Write-Host "  Name:           $($systemInfo.Processor.Name)" -ForegroundColor $script:Colors.Info
        Write-Host "  Cores:          $($systemInfo.Processor.Cores)" -ForegroundColor $script:Colors.Info
        Write-Host "  Logical Procs:  $($systemInfo.Processor.LogicalProcessors)" -ForegroundColor $script:Colors.Info
        Write-Host "  Max Clock:      $($systemInfo.Processor.MaxClockSpeed) MHz" -ForegroundColor $script:Colors.Info
        Write-Host ""
        
        # Storage Information
        Write-Host "STORAGE" -ForegroundColor $script:Colors.Primary
        foreach ($drive in $systemInfo.Storage) {
            Write-Host "  $($drive.Drive) $($drive.VolumeName)" -ForegroundColor $script:Colors.Info
            Write-Host "    Size:         $($drive.Size) GB" -ForegroundColor $script:Colors.Info
            Write-Host "    Free Space:   $($drive.FreeSpace) GB" -ForegroundColor $script:Colors.Info
            
            # Color-code free space percentage
            $percentColor = if ($drive.PercentFree -lt 10) { $script:Colors.Error } 
                            elseif ($drive.PercentFree -lt 20) { $script:Colors.Warning } 
                            else { $script:Colors.Success }
            Write-Host "    Percent Free: $($drive.PercentFree)%" -ForegroundColor $percentColor
            Write-Host ""
        }
        
        # Network Information
        Write-Host "NETWORK ADAPTERS" -ForegroundColor $script:Colors.Primary
        if ($systemInfo.Network.Count -gt 0) {
            foreach ($adapter in $systemInfo.Network) {
                Write-Host "  $($adapter.Name)" -ForegroundColor $script:Colors.Info
                Write-Host "    MAC Address:   $($adapter.MacAddress)" -ForegroundColor $script:Colors.Info
                Write-Host "    Type:          $($adapter.AdapterType)" -ForegroundColor $script:Colors.Info
                if ($adapter.Speed) {
                    $speedInMbps = [math]::Round($adapter.Speed / 1000000, 0)
                    Write-Host "    Speed:         $speedInMbps Mbps" -ForegroundColor $script:Colors.Info
                }
                Write-Host ""
            }
        }
        else {
            Write-Host "  No active network adapters found." -ForegroundColor $script:Colors.Warning
            Write-Host ""
        }
    }
    
    Show-Footer -Options @("Press any key to return to the main menu...")
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    Show-MainMenu
}

function Show-AboutScreen {
    <#
    .SYNOPSIS
        Displays information about the application
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "About AutoByte Professional"
    
    Write-Host "AutoByte Professional v$script:Version" -ForegroundColor $script:Colors.Primary
    Write-Host "Advanced PowerShell Task Automation App with LLM Integration" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    Write-Host "FEATURES" -ForegroundColor $script:Colors.Primary
    Write-Host "• Automated IT task execution from a central script repository" -ForegroundColor $script:Colors.Info
    Write-Host "• AI-powered troubleshooting assistant using Llama LLM" -ForegroundColor $script:Colors.Info
    Write-Host "• Comprehensive system information collection and analysis" -ForegroundColor $script:Colors.Info
    Write-Host "• Category-based script organization for easy navigation" -ForegroundColor $script:Colors.Info
    Write-Host "• Robust error handling and detailed logging" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    Write-Host "SYSTEM REQUIREMENTS" -ForegroundColor $script:Colors.Primary
    Write-Host "• Windows 10/11 or Windows Server 2016+" -ForegroundColor $script:Colors.Info
    Write-Host "• PowerShell 5.1 or higher" -ForegroundColor $script:Colors.Info
    Write-Host "• Internet connection for script downloading" -ForegroundColor $script:Colors.Info
    Write-Host "• Administrator privileges for certain operations" -ForegroundColor $script:Colors.Info
    if (![string]::IsNullOrEmpty($LlamaModelPath)) {
        Write-Host "• Llama LLM model for AI assistant functionality" -ForegroundColor $script:Colors.Info
    }
    Write-Host ""
    
    Write-Host "LOG LOCATION" -ForegroundColor $script:Colors.Primary
    Write-Host "• $script:LogFile" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    Write-Host "ACKNOWLEDGEMENTS" -ForegroundColor $script:Colors.Primary
    Write-Host "• Original scripts by deep1ne8" -ForegroundColor $script:Colors.Info
    Write-Host "• Enhanced interface and LLM integration by Claude" -ForegroundColor $script:Colors.Info
    Write-Host "• Includes Llama LLM technology for AI assistant capabilities" -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    Show-Footer -Options @("Press any key to return to the main menu...")
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    Show-MainMenu
}

function Exit-Application {
    <#
    .SYNOPSIS
        Performs cleanup and exits the application
    #>
    [CmdletBinding()]
    param()
    
    Show-Header -Title "Exiting Application"
    
    Write-Host "Thank you for using $script:AppName!" -ForegroundColor $script:Colors.Success
    Write-Host "Performing cleanup and exiting..." -ForegroundColor $script:Colors.Info
    
    # Perform any necessary cleanup
    Write-Log "$script:AppName v$script:Version terminated" -Level INFO
    
    # Exit with a slight delay for visual feedback
    Start-Sleep -Seconds 1
    return
} 
#endregion

#region Main Execution
try {
    # Initialize environment
    Initialize-Environment
    
    # Check prerequisites
    $isAdmin = Test-AdminPrivileges
    $adminText = if ($isAdmin) { "with" } else { "without" }
    Write-Log "Running $adminText administrator privileges" -Level INFO
    
    # Start main menu
    Show-MainMenu
}
catch {
    # Fatal error handling
    $errorMessage = "Fatal error: $_"
    Write-Log $errorMessage -Level ERROR
    
    Write-Host "`n`n" -NoNewline
    Write-Host "A fatal error has occurred:" -ForegroundColor $script:Colors.Error
    Write-Host $_.Exception.Message -ForegroundColor $script:Colors.Error
    Write-Host "Please check the log file for details: $script:LogFile" -ForegroundColor $script:Colors.Warning
    Write-Host "`n" -NoNewline
    
    # Wait for user acknowledgment
    Write-Host "Press any key to exit..." -ForegroundColor $script:Colors.Warning
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    # Exit with error code
    return
}
#endregion
