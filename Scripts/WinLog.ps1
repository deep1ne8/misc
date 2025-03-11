# PowerShell Log File Interpreter
# Advanced log analysis tool with intuitive UI, categorization, and solution suggestions

# Requires PowerShell 5.1 or higher

using namespace System.Management.Automation.Host
using namespace System.Collections.Generic

#region Configuration
# Define color scheme
$script:colorScheme = @{
    Header       = [ConsoleColor]::Cyan
    MenuSelected = [ConsoleColor]::Black, [ConsoleColor]::Cyan
    MenuNormal   = [ConsoleColor]::White
    Success      = [ConsoleColor]::Green
    Error        = [ConsoleColor]::Red
    Warning      = [ConsoleColor]::Yellow
    Info         = [ConsoleColor]::Blue
    Highlight    = [ConsoleColor]::Magenta
    Default      = [ConsoleColor]::White
}

# Define known log file patterns and solutions database
$script:knownIssues = @{}

# Path for storing learned solutions
$script:solutionsDbPath = Join-Path $env:APPDATA "LogInterpreter\solutions.json"

# Initialize database or load existing one
function Initialize-SolutionsDatabase {
    if (-not (Test-Path (Split-Path $script:solutionsDbPath -Parent))) {
        New-Item -Path (Split-Path $script:solutionsDbPath -Parent) -ItemType Directory -Force | Out-Null
    }
    
    if (Test-Path $script:solutionsDbPath) {
        try {
            $script:knownIssues = Get-Content -Path $script:solutionsDbPath -Raw | ConvertFrom-Json -AsHashtable
        }
        catch {
            Write-Host "Error loading solutions database. Creating a new one." -ForegroundColor $colorScheme.Error
            $script:knownIssues = @{}
            Save-SolutionsDatabase
        }
    }
    else {
        # Initialize with some common patterns
        $script:knownIssues = @{
            "Access denied" = @{
                Pattern = "Access (denied|is denied)"
                Solution = "Check permissions. Run as administrator or modify ACLs for the required resources."
                Occurrences = 0
                Category = "Permissions"
            }
            "Service not started" = @{
                Pattern = "(service|daemon) (failed to start|not started|couldn't start)"
                Solution = "Try starting the service manually via Services.msc or using Start-Service cmdlet."
                Occurrences = 0
                Category = "Services"
            }
            "Port conflict" = @{
                Pattern = "(port|socket) (conflict|already in use|being used)"
                Solution = "Identify which application is using the port with netstat -ano | findstr PORT_NUMBER and either change the port or stop the conflicting application."
                Occurrences = 0
                Category = "Networking"
            }
            "Disk space" = @{
                Pattern = "(disk|space) (full|insufficient|not enough)"
                Solution = "Free up disk space by removing temporary files, using Disk Cleanup, or expanding storage."
                Occurrences = 0
                Category = "Storage"
            }
            "Missing file" = @{
                Pattern = "(file|module) (not found|missing|couldn't be located)"
                Solution = "Verify the file path is correct and the file exists. Reinstall the application if necessary."
                Occurrences = 0
                Category = "Files"
            }
        }
        Save-SolutionsDatabase
    }
}

# Save the solutions database
function Save-SolutionsDatabase {
    $script:knownIssues | ConvertTo-Json -Depth 5 | Set-Content -Path $script:solutionsDbPath -Force
}

# Log file categories with their default locations
$script:logCategories = @{
    "Windows Event Logs" = @{
        Description = "System, Application, Security logs"
        Type = "EventLog"
        Paths = @("System", "Application", "Security", "Setup", "PowerShellCore", "Windows PowerShell")
    }
    "IIS Logs" = @{
        Description = "Internet Information Services logs"
        Type = "File"
        Paths = @("C:\inetpub\logs\LogFiles")
    }
    "Custom Application Logs" = @{
        Description = "Various application log files"
        Type = "File"
        Paths = @("C:\ProgramData", "C:\Program Files", "C:\Program Files (x86)")
        Extensions = @(".log", ".txt", ".log1", ".log2")
    }
    "PowerShell Transcripts" = @{
        Description = "PowerShell transcript files"
        Type = "File"
        Paths = @(Join-Path $env:USERPROFILE "Documents\PowerShell\Transcripts")
        Extensions = @(".txt")
    }
    "Windows Update Logs" = @{
        Description = "Windows Update service logs"
        Type = "File"
        Paths = @("$env:SystemRoot\Logs\WindowsUpdate")
        Extensions = @(".log", ".etl")
    }
    "SCCM/ConfigMgr Logs" = @{
        Description = "Configuration Manager client logs"
        Type = "File"
        Paths = @("$env:SystemRoot\CCM\Logs", "$env:SystemRoot\SysWOW64\CCM\Logs")
        Extensions = @(".log")
    }
}
#endregion

#region UI Functions
# Function to draw a centered header
function Write-Header {
    param (
        [string]$Text
    )
    
    $width = $host.UI.RawUI.WindowSize.Width
    $padding = [math]::Max(0, ($width - $Text.Length - 4) / 2)
    
    Write-Host (" " * [math]::Floor($padding)) -NoNewline
    Write-Host "╔" -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host "═" * ($Text.Length + 2) -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host "╗" -ForegroundColor $colorScheme.Header
    
    Write-Host (" " * [math]::Floor($padding)) -NoNewline
    Write-Host "║ " -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host $Text -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host " ║" -ForegroundColor $colorScheme.Header
    
    Write-Host (" " * [math]::Floor($padding)) -NoNewline
    Write-Host "╚" -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host "═" * ($Text.Length + 2) -NoNewline -ForegroundColor $colorScheme.Header
    Write-Host "╝" -ForegroundColor $colorScheme.Header
    
    Write-Host ""
}

# Function to display a menu and get user selection
function Show-Menu {
    param (
        [string]$Title,
        [string[]]$Options,
        [int]$DefaultSelection = 0
    )
    
    $currentSelection = $DefaultSelection
    $startY = $host.UI.RawUI.CursorPosition.Y
    $width = $host.UI.RawUI.WindowSize.Width
    
    Write-Header $Title
    
    $maxOptionLength = ($Options | Measure-Object -Maximum -Property Length).Maximum
    $menuWidth = $maxOptionLength + 6  # 2 spaces + 2 brackets + 2 padding
    
    # Render the menu
    function Render-Menu {
        $cursorPosition = $host.UI.RawUI.CursorPosition
        $cursorPosition.Y = $startY + 2
        $host.UI.RawUI.CursorPosition = $cursorPosition
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            $padding = [math]::Max(0, ($width - $menuWidth) / 2)
            Write-Host (" " * [math]::Floor($padding)) -NoNewline
            
            if ($i -eq $currentSelection) {
                Write-Host "►" -NoNewline -ForegroundColor $colorScheme.MenuSelected[0]
                Write-Host " [" -NoNewline -ForegroundColor $colorScheme.MenuSelected[0]
                Write-Host $Options[$i].PadRight($maxOptionLength) -NoNewline -ForegroundColor $colorScheme.MenuSelected[0] -BackgroundColor $colorScheme.MenuSelected[1]
                Write-Host "] " -ForegroundColor $colorScheme.MenuSelected[0]
            }
            else {
                Write-Host "  [" -NoNewline -ForegroundColor $colorScheme.MenuNormal
                Write-Host $Options[$i].PadRight($maxOptionLength) -NoNewline -ForegroundColor $colorScheme.MenuNormal
                Write-Host "]  " -ForegroundColor $colorScheme.MenuNormal
            }
        }
        
        Write-Host ""
        Write-Host "Use ↑/↓ to navigate, Enter to select" -ForegroundColor $colorScheme.Info
    }
    
    Render-Menu
    
    # Handle keyboard input
    while ($true) {
        $keyInfo = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($keyInfo.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentSelection -gt 0) {
                    $currentSelection--
                    Render-Menu
                }
                break
            }
            40 { # Down arrow
                if ($currentSelection -lt $Options.Count - 1) {
                    $currentSelection++
                    Render-Menu
                }
                break
            }
            13 { # Enter
                return $currentSelection
                break
            }
        }
    }
}

# Function to display an interactive list with scrolling
function Show-ScrollableList {
    param (
        [string]$Title,
        [PSObject[]]$Items,
        [scriptblock]$DisplayScript,
        [int]$PageSize = 10
    )
    
    if ($Items.Count -eq 0) {
        Write-Host "No items to display." -ForegroundColor $colorScheme.Warning
        Read-Host "Press Enter to continue"
        return $null
    }
    
    $currentSelection = 0
    $currentPage = 0
    $totalPages = [math]::Ceiling($Items.Count / $PageSize)
    
    while ($true) {
        Clear-Host
        Write-Header $Title
        
        $start = $currentPage * $PageSize
        $end = [math]::Min($start + $PageSize - 1, $Items.Count - 1)
        
        Write-Host "Page $($currentPage + 1) of $totalPages" -ForegroundColor $colorScheme.Info
        Write-Host "──────────────────────────────────" -ForegroundColor $colorScheme.Info
        
        for ($i = $start; $i -le $end; $i++) {
            if ($i -eq $currentSelection) {
                Write-Host "►" -NoNewline -ForegroundColor $colorScheme.MenuSelected[0]
                
                # Call the display script with the current item
                $displayText = & $DisplayScript $Items[$i]
                
                # Write the display text with selected formatting
                Write-Host " $displayText" -ForegroundColor $colorScheme.MenuSelected[0] -BackgroundColor $colorScheme.MenuSelected[1]
            }
            else {
                Write-Host "  " -NoNewline
                
                # Call the display script with the current item
                $displayText = & $DisplayScript $Items[$i]
                
                Write-Host $displayText -ForegroundColor $colorScheme.MenuNormal
            }
        }
        
        Write-Host ""
        Write-Host "Use ↑/↓ to navigate, Enter to select, PageUp/PageDown to change pages, Esc to return" -ForegroundColor $colorScheme.Info
        
        $keyInfo = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($keyInfo.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentSelection -gt $start) {
                    $currentSelection--
                }
                elseif ($currentPage -gt 0) {
                    $currentPage--
                    $currentSelection = $currentPage * $PageSize + $PageSize - 1
                }
                break
            }
            40 { # Down arrow
                if ($currentSelection -lt $end) {
                    $currentSelection++
                }
                elseif ($currentPage -lt $totalPages - 1) {
                    $currentPage++
                    $currentSelection = $currentPage * $PageSize
                }
                break
            }
            33 { # Page Up
                if ($currentPage -gt 0) {
                    $currentPage--
                    $currentSelection = $currentPage * $PageSize
                }
                break
            }
            34 { # Page Down
                if ($currentPage -lt $totalPages - 1) {
                    $currentPage++
                    $currentSelection = $currentPage * $PageSize
                }
                break
            }
            13 { # Enter
                return $Items[$currentSelection]
                break
            }
            27 { # Escape
                return $null
                break
            }
        }
    }
}

# Function to display search and filter UI
function Show-SearchUI {
    param (
        [string]$Title,
        [string]$DefaultSearch = ""
    )
    
    Clear-Host
    Write-Header $Title
    
    Write-Host "Enter search term (wildcards * and ? supported): " -ForegroundColor $colorScheme.Info -NoNewline
    $searchTerm = if ($DefaultSearch) { 
        Write-Host $DefaultSearch -ForegroundColor $colorScheme.Highlight
        $DefaultSearch 
    } else { 
        $Host.UI.ReadLine() 
    }
    
    return $searchTerm
}
#endregion

#region Log Collection and Analysis Functions
# Function to get log files for a specific category
function Get-LogFiles {
    param (
        [string]$Category
    )
    
    $categoryInfo = $script:logCategories[$Category]
    $logFiles = @()
    
    if ($categoryInfo.Type -eq "EventLog") {
        $logFiles = Get-WinEvent -ListLog $categoryInfo.Paths -ErrorAction SilentlyContinue | 
                     Where-Object { $_.RecordCount -gt 0 } | 
                     Select-Object LogName, RecordCount, FileSize, LastWriteTime, LogFilePath
    }
    elseif ($categoryInfo.Type -eq "File") {
        foreach ($path in $categoryInfo.Paths) {
            if (Test-Path $path) {
                $extensions = if ($categoryInfo.ContainsKey("Extensions")) { $categoryInfo.Extensions } else { @(".log", ".txt") }
                
                $files = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue | 
                          Where-Object { $extensions -contains $_.Extension } | 
                          Select-Object FullName, Name, Length, LastWriteTime
                
                $logFiles += $files
            }
        }
    }
    
    return $logFiles
}

# Function to extract log entries from a file
function Get-LogEntries {
    param (
        [PSObject]$LogFile,
        [string]$Category,
        [int]$MaxEntries = 1000,
        [string]$SearchTerm = "",
        [DateTime]$StartTime = [DateTime]::MinValue,
        [DateTime]$EndTime = [DateTime]::MaxValue
    )
    
    $categoryInfo = $script:logCategories[$Category]
    $entries = @()
    
    try {
        if ($categoryInfo.Type -eq "EventLog") {
            $filter = @{
                LogName = $LogFile.LogName
                StartTime = $StartTime
                EndTime = $EndTime
            }
            
            if ($SearchTerm) {
                $events = Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEntries -ErrorAction SilentlyContinue | 
                           Where-Object { $_.Message -like "*$SearchTerm*" -or $_.ProviderName -like "*$SearchTerm*" }
            }
            else {
                $events = Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEntries -ErrorAction SilentlyContinue
            }
            
            $entries = $events | Select-Object TimeCreated, LevelDisplayName, ProviderName, Id, Message
        }
        elseif ($categoryInfo.Type -eq "File") {
            $content = Get-Content -Path $LogFile.FullName -ErrorAction Stop
            
            # Try to determine log format
            $logFormat = Detect-LogFormat -LogContent $content -LogFileName $LogFile.Name
            
            $parsedEntries = Parse-LogContent -LogContent $content -LogFormat $logFormat
            
            if ($SearchTerm) {
                $parsedEntries = $parsedEntries | Where-Object { 
                    ($_.Timestamp -ge $StartTime -and $_.Timestamp -le $EndTime) -and
                    ($_.Message -like "*$SearchTerm*" -or $_.Level -like "*$SearchTerm*" -or $_.Source -like "*$SearchTerm*")
                }
            }
            else {
                $parsedEntries = $parsedEntries | Where-Object { 
                    $_.Timestamp -ge $StartTime -and $_.Timestamp -le $EndTime
                }
            }
            
            $entries = $parsedEntries | Select-Object -First $MaxEntries
        }
    }
    catch {
        Write-Host "Error reading log file: $_" -ForegroundColor $colorScheme.Error
    }
    
    return $entries
}

# Function to detect log file format
function Detect-LogFormat {
    param (
        [string[]]$LogContent,
        [string]$LogFileName
    )
    
    # Default format if we can't detect
    $format = @{
        Type = "Unknown"
        TimestampFormat = $null
        RegexPattern = $null
    }
    
    # Return early if content is empty
    if ($null -eq $LogContent -or $LogContent.Count -eq 0) {
        return $format
    }
    
    # Sample the first few lines
    $sampleSize = [Math]::Min(10, $LogContent.Count)
    $sampleLines = $LogContent | Select-Object -First $sampleSize
    
    # Check for W3C format (IIS logs)
    if ($LogFileName -like "*iis*" -or $sampleLines -join "`n" -match "#Fields: date time") {
        $format.Type = "W3C"
        $format.TimestampFormat = "yyyy-MM-dd HH:mm:ss"
        return $format
    }
    
    # Check for common timestamp patterns
    $timestampPatterns = @(
        # ISO 8601
        @{
            Regex = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?([+-]\d{2}:\d{2}|Z)?'
            Format = if ($matches -and $matches[2]) { "yyyy-MM-ddTHH:mm:ss.fffK" } else { "yyyy-MM-ddTHH:mm:ss.fff" }
        },
        # Common log format with date and time
        @{
            Regex = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(,\d+)?'
            Format = "yyyy-MM-dd HH:mm:ss,fff"
        },
        # Windows event log format
        @{
            Regex = '^Log Name:'
            Format = "Windows Event Log"
        },
        # Common time format without date
        @{
            Regex = '^\[\d{2}:\d{2}:\d{2}\]'
            Format = "HH:mm:ss"
        }
    )
    
    foreach ($line in $sampleLines) {
        foreach ($pattern in $timestampPatterns) {
            if ($line -match $pattern.Regex) {
                $format.Type = "Structured"
                $format.TimestampFormat = $pattern.Format
                return $format
            }
        }
    }
    
    # If we can't find a standard format, try to identify common log levels
    $logLevelPattern = '\b(ERROR|WARN|INFO|DEBUG|FATAL|TRACE|CRITICAL|SEVERE)\b'
    
    if (($sampleLines -join "`n") -match $logLevelPattern) {
        $format.Type = "Structured"
        return $format
    }
    
    return $format
}

# Function to parse log content based on detected format
function Parse-LogContent {
    param (
        [string[]]$LogContent,
        [hashtable]$LogFormat
    )
    
    $entries = @()
    
    # Return if content is empty
    if ($null -eq $LogContent -or $LogContent.Count -eq 0) {
        return $entries
    }

    switch ($LogFormat.Type) {
        "W3C" {
            # Parse W3C format (typically IIS logs)
            $headerLine = $LogContent | Where-Object { $_ -like "#Fields:*" } | Select-Object -First 1
            
            if ($headerLine) {
                $fields = ($headerLine -replace "#Fields: ", "") -split " "
                $dataLines = $LogContent | Where-Object { -not $_.StartsWith("#") -and $_.Trim() -ne "" }
                
                foreach ($line in $dataLines) {
                    $values = $line -split " "
                    
                    if ($values.Count -ge 2) {
                        $timestamp = if ($fields -contains "date" -and $fields -contains "time") {
                            $dateIndex = [array]::IndexOf($fields, "date")
                            $timeIndex = [array]::IndexOf($fields, "time")
                            
                            if ($dateIndex -ge 0 -and $timeIndex -ge 0 -and $values.Count -gt [math]::Max($dateIndex, $timeIndex)) {
                                try {
                                    [DateTime]::ParseExact("$($values[$dateIndex]) $($values[$timeIndex])", "yyyy-MM-dd HH:mm:ss", $null)
                                }
                                catch {
                                    Get-Date
                                }
                            }
                            else {
                                Get-Date
                            }
                        }
                        else {
                            Get-Date
                        }
                        
                        $entry = [PSCustomObject]@{
                            Timestamp = $timestamp
                            Level = "INFO"
                            Source = "W3C"
                            Message = $line
                            RawData = $values
                        }
                        
                        $entries += $entry
                    }
                }
            }
        }
        "Structured" {
            # Attempt to parse structured logs with timestamps and possibly log levels
            $currentEntry = $null
            $entryLines = @()
            
            foreach ($line in $LogContent) {
                $isNewEntry = $false
                
                # Try to detect timestamp at the beginning of the line
                if ($LogFormat.TimestampFormat -eq "Windows Event Log") {
                    # Windows Event Log syntax - special handling
                    if ($line -match "^Log Name:") {
                        $isNewEntry = $true
                    }
                }
                elseif ($line -match '^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}') {
                    # ISO-like timestamp
                    $isNewEntry = $true
                }
                elseif ($line -match '^\[\d{2}:\d{2}:\d{2}\]') {
                    # Time in brackets
                    $isNewEntry = $true
                }
                
                # If this is a new entry and we have data for the previous entry, add it
                if ($isNewEntry -and $entryLines.Count -gt 0) {
                    $entryText = $entryLines -join "`n"
                    
                    # Extract log level if present
                    $logLevel = "INFO"
                    if ($entryText -match '\b(ERROR|WARN|INFO|DEBUG|FATAL|TRACE|CRITICAL|SEVERE)\b') {
                        $logLevel = $matches[1]
                    }
                    
                    # Extract timestamp if present
                    $timestamp = $null
                    $timestampRegex = @(
                        '(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(\.\d+)?)',
                        '\[(\d{2}:\d{2}:\d{2})\]'
                    )
                    
                    foreach ($regex in $timestampRegex) {
                        if ($entryText -match $regex) {
                            try {
                                $timestampStr = $matches[1]
                                $timestamp = [DateTime]::Parse($timestampStr)
                                break
                            }
                            catch {
                                # Continue to next regex if parsing fails
                            }
                        }
                    }
                    
                    if ($null -eq $timestamp) {
                        $timestamp = Get-Date
                    }
                    
                    # Create log entry
                    $entry = [PSCustomObject]@{
                        Timestamp = $timestamp
                        Level = $logLevel
                        Source = "LOG"
                        Message = $entryText
                        RawData = $entryLines
                    }
                    
                    $entries += $entry
                    $entryLines = @()
                }
                
                $entryLines += $line
            }
            
            # Process the last entry if there is one
            if ($entryLines.Count -gt 0) {
                $entryText = $entryLines -join "`n"
                
                # Extract log level if present
                $logLevel = "INFO"
                if ($entryText -match '\b(ERROR|WARN|INFO|DEBUG|FATAL|TRACE|CRITICAL|SEVERE)\b') {
                    $logLevel = $matches[1]
                }
                
                # Extract timestamp if present
                $timestamp = $null
                $timestampRegex = @(
                    '(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(\.\d+)?)',
                    '\[(\d{2}:\d{2}:\d{2})\]'
                )
                
                foreach ($regex in $timestampRegex) {
                    if ($entryText -match $regex) {
                        try {
                            $timestampStr = $matches[1]
                            $timestamp = [DateTime]::Parse($timestampStr)
                            break
                        }
                        catch {
                            # Continue to next regex if parsing fails
                        }
                    }
                }
                
                if ($null -eq $timestamp) {
                    $timestamp = Get-Date
                }
                
                # Create log entry
                $entry = [PSCustomObject]@{
                    Timestamp = $timestamp
                    Level = $logLevel
                    Source = "LOG"
                    Message = $entryText
                    RawData = $entryLines
                }
                
                $entries += $entry
            }
        }
        default {
            # Unstructured log - treat each line as a separate entry
            foreach ($line in $LogContent) {
                if ($line.Trim() -ne "") {
                    $entry = [PSCustomObject]@{
                        Timestamp = Get-Date
                        Level = "INFO"
                        Source = "LOG"
                        Message = $line
                        RawData = $line
                    }
                    
                    $entries += $entry
                }
            }
        }
    }
    
    return $entries
}

# Function to analyze log entries for known issues
function Analyze-LogEntries {
    param (
        [PSObject[]]$Entries
    )
    
    $results = @()
    
    foreach ($issue in $script:knownIssues.GetEnumerator()) {
        $pattern = $issue.Value.Pattern
        $matchingEntries = $Entries | Where-Object { $_.Message -match $pattern }
        
        if ($matchingEntries.Count -gt 0) {
            $result = [PSCustomObject]@{
                IssueName = $issue.Key
                Pattern = $pattern
                Solution = $issue.Value.Solution
                Category = $issue.Value.Category
                MatchingEntries = $matchingEntries
                EntryCount = $matchingEntries.Count
            }
            
            $results += $result
            
            # Update occurrence count
            $script:knownIssues[$issue.Key].Occurrences += 1
        }
    }
    
    # Sort by entry count descending
    $results = $results | Sort-Object -Property EntryCount -Descending
    
    return $results
}

# Function to learn a new solution
function Add-LearnedSolution {
    param (
        [string]$IssueName,
        [string]$Pattern,
        [string]$Solution,
        [string]$Category
    )
    
    if (-not $script:knownIssues.ContainsKey($IssueName)) {
        $script:knownIssues[$IssueName] = @{
            Pattern = $Pattern
            Solution = $Solution
            Occurrences = 1
            Category = $Category
        }
        
        Save-SolutionsDatabase
        return $true
    }
    
    return $false
}
#endregion

#region Main Application Functions
# Function to display log entries with colorized output
function Show-LogEntries {
    param (
        [PSObject]$LogFile,
        [string]$Category,
        [PSObject[]]$Entries
    )
    
    if ($null -eq $Entries -or $Entries.Count -eq 0) {
        Write-Host "No log entries found matching the criteria." -ForegroundColor $colorScheme.Warning
        return
    }
    
    $selectedEntry = Show-ScrollableList -Title "Log Entries" -Items $Entries -DisplayScript {
        param($entry)
        
        # Format timestamp
        $timeStr = $entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
        
        # Format level with color
        $levelStr = switch -Regex ($entry.Level) {
            "ERROR|FATAL|CRITICAL|SEVERE" { "Error" }
            "WARN" { "Warning" }
            "INFO" { "Info" }
            default { $entry.Level }
        }
        
        # Get first line of message
        $messageLine = ($entry.Message -split '\r?\n')[0]
        if ($messageLine.Length -gt 80) {
            $messageLine = $messageLine.Substring(0, 77) + "..."
        }
        
        return "[$timeStr] $levelStr: $messageLine"
    }
    
    if ($null -ne $selectedEntry) {
        Show-LogEntryDetails -Entry $selectedEntry
    }
    
    # After viewing entries, offer to analyze them
    Clear-Host
    Write-Header "Analyze Log Entries"
    
    Write-Host "Would you like to analyze these log entries for known issues? (y/n)" -ForegroundColor $colorScheme.Question
    $response = $Host.UI.ReadLine()
    
    if ($response -eq "y") {
        $analysisResults = Analyze-LogEntries -Entries $Entries
        Show-AnalysisResults -Results $analysisResults
    }
}

# Function to show analysis results with suggested solutions
function Show-AnalysisResults {
    param (
        [PSObject[]]$Results
    )
    
    Clear-Host
    Write-Header "Log Analysis Results"
    
    if ($Results.Count -eq 0) {
        Write-Host "No known issues were found in the selected log entries." -ForegroundColor $colorScheme.Info
        Write-Host "Would you like to add a custom solution for a pattern in these logs? (y/n)" -ForegroundColor $colorScheme.Question
        
        $response = $Host.UI.ReadLine()
        if ($response -eq "y") {
            Add-CustomSolution
        }
        
        return
    }
    
    Write-Host "Found $($Results.Count) potential issues in the log entries:" -ForegroundColor $colorScheme.Info
    Write-Host ""
    
    foreach ($result in $Results) {
        Write-Host "Issue: " -NoNewline -ForegroundColor $colorScheme.Highlight
        Write-Host $result.IssueName -ForegroundColor $colorScheme.Default
        
        Write-Host "Category: " -NoNewline -ForegroundColor $colorScheme.Info
        Write-Host $result.Category -ForegroundColor $colorScheme.Default
        
        Write-Host "Matches: " -NoNewline -ForegroundColor $colorScheme.Info
        Write-Host "$($result.EntryCount) entries" -ForegroundColor $colorScheme.Default
        
        Write-Host "Suggested Solution: " -ForegroundColor $colorScheme.Success
        Write-Host "  $($result.Solution)" -ForegroundColor $colorScheme.Default
        
        Write-Host "───────────────────────────────────────────────" -ForegroundColor $colorScheme.Info
    }
    
    Write-Host ""
    Write-Host "Would you like to view matching entries for a specific issue? (y/n)" -ForegroundColor $colorScheme.Question
    
    $response = $Host.UI.ReadLine()
    if ($response -eq "y") {
        $selectedIssue = Show-ScrollableList -Title "Select Issue" -Items $Results -DisplayScript {
            param($item)
            return "$($item.IssueName) ($($item.EntryCount) entries)"
        }
        
        if ($null -ne $selectedIssue) {
            $selectedEntry = Show-ScrollableList -Title "Matching Entries for '$($selectedIssue.IssueName)'" -Items $selectedIssue.MatchingEntries -DisplayScript {
                param($entry)
                $timeStr = $entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
                return "[$timeStr] $($entry.Level): $(($entry.Message -split '\r?\n')[0])"
            }
            
            if ($null -ne $selectedEntry) {
                Show-LogEntryDetails -Entry $selectedEntry
            }
        }
    }
    
    Write-Host ""
    Write-Host "Would you like to add a custom solution for a pattern that wasn't detected? (y/n)" -ForegroundColor $colorScheme.Question
    
    $response = $Host.UI.ReadLine()
    if ($response -eq "y") {
        Add-CustomSolution
    }
}

# Function to manage log filtering options
function Show-FilterOptions {
    param (
        [PSObject]$LogFile,
        [string]$Category,
        [DateTime]$StartTime = [DateTime]::Now.AddDays(-1),
        [DateTime]$EndTime = [DateTime]::Now,
        [string]$SearchTerm = ""
    )
    
    $options = @(
        "Change Start Time (Current: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss')))",
        "Change End Time (Current: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))",
        "Set Search Term (Current: $(if ($SearchTerm) { $SearchTerm } else { 'None' }))",
        "Apply Filters and View Logs",
        "Return to Previous Menu"
    )
    
    while ($true) {
        Clear-Host
        $selection = Show-Menu -Title "Log Filter Options" -Options $options
        
        switch ($selection) {
            0 { # Change Start Time
                Clear-Host
                Write-Header "Set Start Time"
                Write-Host "Enter a new start time (yyyy-MM-dd HH:mm:ss) or relative time (e.g., -1d, -12h):" -ForegroundColor $colorScheme.Info
                $input = $Host.UI.ReadLine()
                
                if ($input -match '^-(\d+)([dhm])$') {
                    $value = [int]$matches[1]
                    $unit = $matches[2]
                    
                    $StartTime = switch ($unit) {
                        'd' { [DateTime]::Now.AddDays(-$value) }
                        'h' { [DateTime]::Now.AddHours(-$value) }
                        'm' { [DateTime]::Now.AddMinutes(-$value) }
                    }
                }
                elseif ($input) {
                    try {
                        $StartTime = [DateTime]::Parse($input)
                    }
                    catch {
                        Write-Host "Invalid date format. Please try again." -ForegroundColor $colorScheme.Error
                        Start-Sleep -Seconds 2
                    }
                }
                
                $options[0] = "Change Start Time (Current: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss')))"
            }
            1 { # Change End Time
                Clear-Host
                Write-Header "Set End Time"
                Write-Host "Enter a new end time (yyyy-MM-dd HH:mm:ss) or 'now' for current time:" -ForegroundColor $colorScheme.Info
                $input = $Host.UI.ReadLine()
                
                if ($input -eq "now") {
                    $EndTime = [DateTime]::Now
                }
                elseif ($input) {
                    try {
                        $EndTime = [DateTime]::Parse($input)
                    }
                    catch {
                        Write-Host "Invalid date format. Please try again." -ForegroundColor $colorScheme.Error
                        Start-Sleep -Seconds 2
                    }
                }
                
                $options[1] = "Change End Time (Current: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))"
            }
            2 { # Set Search Term
                $newTerm = Show-SearchUI -Title "Set Search Term" -DefaultSearch $SearchTerm
                if ($null -ne $newTerm) {
                    $SearchTerm = $newTerm
                    $options[2] = "Set Search Term (Current: $(if ($SearchTerm) { $SearchTerm } else { 'None' }))"
                }
            }
            3 { # Apply and View
                $entries = Get-LogEntries -LogFile $LogFile -Category $Category -StartTime $StartTime -EndTime $EndTime -SearchTerm $SearchTerm
                Show-LogEntries -LogFile $LogFile -Category $Category -Entries $entries
            }
            4 { # Return
                return
            }
        }
    }
}

# Function to add a custom solution
function Add-CustomSolution {
    Clear-Host
    Write-Header "Add Custom Solution"
    
    Write-Host "Provide a name for this issue:" -ForegroundColor $colorScheme.Info
    $issueName = $Host.UI.ReadLine()
    
    Write-Host "Enter a regular expression pattern to match in log entries:" -ForegroundColor $colorScheme.Info
    $pattern = $Host.UI.ReadLine()
    
    Write-Host "Enter a category for this issue:" -ForegroundColor $colorScheme.Info
    $category = $Host.UI.ReadLine()
    
    Write-Host "Provide a solution for this issue:" -ForegroundColor $colorScheme.Info
    $solution = $Host.UI.ReadLine()
    
    if ($issueName -and $pattern -and $solution) {
        $added = Add-LearnedSolution -IssueName $issueName -Pattern $pattern -Solution $solution -Category $category
        
        if ($added) {
            Write-Host "Solution added successfully! It will be used for future log analysis." -ForegroundColor $colorScheme.Success
        }
        else {
            Write-Host "A solution with this name already exists. Please try again with a different name." -ForegroundColor $colorScheme.Warning
        }
    }
    else {
        Write-Host "All fields are required. Solution was not added." -ForegroundColor $colorScheme.Error
    }
    
    Write-Host ""
    Write-Host "Press Enter to continue..." -ForegroundColor $colorScheme.Info
    [void][Console]::ReadKey($true)
}
# Function to display log entry details with syntax highlighting
function Show-LogEntryDetails {
    param (
        [PSObject]$Entry
    )
    
    Clear-Host
    Write-Header "Log Entry Details"
    
    # Display metadata
    Write-Host "Timestamp: " -NoNewline -ForegroundColor $colorScheme.Info
    Write-Host $Entry.Timestamp -ForegroundColor $colorScheme.Default
    
    Write-Host "Level:     " -NoNewline -ForegroundColor $colorScheme.Info
    
    # Color-code the level
    switch -Regex ($Entry.Level) {
        "ERROR|FATAL|CRITICAL|SEVERE" {
            Write-Host $Entry.Level -ForegroundColor $colorScheme.Error
        }
        "WARN" {
            Write-Host $Entry.Level -ForegroundColor $colorScheme.Warning
        }
        "INFO" {
            Write-Host $Entry.Level -ForegroundColor $colorScheme.Success
        }
        default {
            Write-Host $Entry.Level -ForegroundColor $colorScheme.Default
        }
    }
    
    Write-Host "Source:    " -NoNewline -ForegroundColor $colorScheme.Info
    Write-Host $Entry.Source -ForegroundColor $colorScheme.Default
    
    Write-Host ""
    Write-Host "Message:" -ForegroundColor $colorScheme.Info
    
    # Syntax highlight the message
    $message = $Entry.Message
    
    # Highlight patterns in the message
    try {
        if ($null -ne $message) {
            $message = $message -replace '(error|exception|failed|failure)', 
                                         { param($m) "`e[91m$($m.Value)`e[0m" }  # Errors in red
            $message = $message -replace '(warning|warn)', 
                                         { param($m) "`e[93m$($m.Value)`e[0m" }  # Warnings in yellow
            $message = $message -replace '(success|succeeded|completed)', 
                                         { param($m) "`e[92m$($m.Value)`e[0m" }  # Success in green
            $message = $message -replace '(\d+\.\d+\.\d+\.\d+)', 
                                         { param($m) "`e[96m$($m.Value)`e[0m" }  # IP addresses in cyan
            $message = $message -replace '("[^"]*")', 
                                         { param($m) { "`e[95m$($m.Value)`e[0m" } }  # Quoted strings in magenta

            Write-Host $message
        }
        else {
            Write-Host "Could not display log entry message due to null value." -ForegroundColor $colorScheme.Error
        }
    }
    catch {
        Write-Host "An error occurred while processing the message: $($_.Exception.Message)" -ForegroundColor $colorScheme.Error
    }

    Write-Host ""
    Write-Host "Press Enter to return..." -ForegroundColor $colorScheme.Info

    try {
        while (-not $Host.UI.RawUI.KeyAvailable) {
            Start-Sleep -Milliseconds 100
        }
        $null = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown)
    }
    catch {
        Write-Host "Failed to read key press. Please try again. Error: $($_.Exception.Message)" -ForegroundColor $colorScheme.Error
    }
}

# Main function to run the log interpreter
function Start-LogInterpreter {
    # Initialize the solutions database
    Initialize-SolutionsDatabase
    
    while ($true) {
        Clear-Host
        
        # Display main menu
        $mainOptions = @(
            "Select Log Category",
            "Search Across All Logs",
            "View Known Solutions Database",
            "Add Custom Solution",
            "Exit"
        )
        
        $mainSelection = Show-Menu -Title "PowerShell Log Interpreter" -Options $mainOptions
        
        switch ($mainSelection) {
            0 { # Select Log Category
                $categories = $script:logCategories.Keys | Sort-Object
                $categorySelection = Show-Menu -Title "Select Log Category" -Options $categories
                
                if ($categorySelection -lt $categories.Count) {
                    $selectedCategory = $categories[$categorySelection]
                    $logFiles = Get-LogFiles -Category $selectedCategory
                    
                    if ($logFiles.Count -eq 0) {
                        Clear-Host
                        Write-Host "No log files found for category: $selectedCategory" -ForegroundColor $colorScheme.Warning
                        Write-Host "Press Enter to continue..." -ForegroundColor $colorScheme.Info
                        [void][Console]::ReadKey($true)
                        continue
                    }
                    
                    $selectedLogFile = Show-ScrollableList -Title "Select Log File" -Items $logFiles -DisplayScript {
                        param($item)
                        
                        if ($item.PSObject.Properties.Name -contains "LogName") {
                            # Event log format
                            return "$($item.LogName) - $($item.RecordCount) records"
                        }
                        else {
                            # File format
                            $size = if ($item.Length -gt 1MB) {
                                "{0:N2} MB" -f ($item.Length / 1MB)
                            }
                            else {
                                "{0:N2} KB" -f ($item.Length / 1KB)
                            }
                            
                            return "$($item.Name) - $size - Last modified: $($item.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))"
                        }
                    }
                    
                    if ($null -ne $selectedLogFile) {
                        # Show filter options for the selected log file
                        Show-FilterOptions -LogFile $selectedLogFile -Category $selectedCategory
                    }
                }
            }
            1 { # Search Across All Logs
                $searchTerm = Show-SearchUI -Title "Search Across All Logs"
                
                if ($searchTerm) {
                    Clear-Host
                    Write-Header "Searching for '$searchTerm' across all logs"
                    
                    $allResults = @()
                    $categories = $script:logCategories.Keys | Sort-Object
                    
                    foreach ($category in $categories) {
                        Write-Host "Searching in $category..." -ForegroundColor $colorScheme.Info
                        
                        $logFiles = Get-LogFiles -Category $category
                        $resultCount = 0
                        
                        foreach ($logFile in $logFiles) {
                            $entries = Get-LogEntries -LogFile $logFile -Category $category -SearchTerm $searchTerm -MaxEntries 100
                            
                            if ($entries.Count -gt 0) {
                                $result = [PSCustomObject]@{
                                    Category = $category
                                    LogFile = $logFile
                                    EntryCount = $entries.Count
                                    Entries = $entries
                                }
                                
                                $allResults += $result
                                $resultCount += $entries.Count
                            }
                        }
                        
                        Write-Host "  Found $resultCount matches" -ForegroundColor (if ($resultCount -gt 0) { $colorScheme.Success } else { $colorScheme.Default })
                    }
                    
                    if ($allResults.Count -gt 0) {
                        Write-Host ""
                        Write-Host "Found matches in $($allResults.Count) log files. Press Enter to view results..." -ForegroundColor $colorScheme.Success
                        [void][Console]::ReadKey($true)
                        
                        $selectedResult = Show-ScrollableList -Title "Search Results for '$searchTerm'" -Items $allResults -DisplayScript {
                            param($item)
                            
                            $logName = if ($item.LogFile.PSObject.Properties.Name -contains "LogName") {
                                $item.LogFile.LogName
                            }
                            else {
                                $item.LogFile.Name
                            }
                            
                            return "$($item.Category) - $logName ($($item.EntryCount) matches)"
                        }
                        
                        if ($null -ne $selectedResult) {
                            Show-LogEntries -LogFile $selectedResult.LogFile -Category $selectedResult.Category -Entries $selectedResult.Entries
                        }
                    }
                    else {
                        Write-Host ""
                        Write-Host "No matches found for '$searchTerm'. Press Enter to continue..." -ForegroundColor $colorScheme.Warning
                        [void][Console]::ReadKey($true)
                    }
                }
            }
            2 { # View Known Solutions Database
                Clear-Host
                Write-Header "Known Solutions Database"
                
                $solutions = $script:knownIssues.GetEnumerator() | ForEach-Object {
                    [PSCustomObject]@{
                        Name = $_.Key
                        Pattern = $_.Value.Pattern
                        Solution = $_.Value.Solution
                        Category = $_.Value.Category
                        Occurrences = $_.Value.Occurrences
                    }
                } | Sort-Object -Property Occurrences -Descending
                
                if ($solutions.Count -eq 0) {
                    Write-Host "No solutions in database yet. Add custom solutions to build your knowledge base." -ForegroundColor $colorScheme.Info
                    Write-Host "Press Enter to continue..." -ForegroundColor $colorScheme.Info
                    [void][Console]::ReadKey($true)
                    continue
                }
                
                $selectedSolution = Show-ScrollableList -Title "Known Solutions" -Items $solutions -DisplayScript {
                    param($item)
                    return "$($item.Name) - $($item.Category) (Used $($item.Occurrences) times)"
                }
                
                if ($null -ne $selectedSolution) {
                    Clear-Host
                    Write-Header "Solution Details"
                    
                    Write-Host "Name: " -NoNewline -ForegroundColor $colorScheme.Info
                    Write-Host $selectedSolution.Name -ForegroundColor $colorScheme.Default
                    
                    Write-Host "Category: " -NoNewline -ForegroundColor $colorScheme.Info
                    Write-Host $selectedSolution.Category -ForegroundColor $colorScheme.Default
                    
                    Write-Host "Pattern: " -NoNewline -ForegroundColor $colorScheme.Info
                    Write-Host $selectedSolution.Pattern -ForegroundColor $colorScheme.Default
                    
                    Write-Host "Usage Count: " -NoNewline -ForegroundColor $colorScheme.Info
                    Write-Host $selectedSolution.Occurrences -ForegroundColor $colorScheme.Default
                    
                    Write-Host ""
                    Write-Host "Solution:" -ForegroundColor $colorScheme.Success
                    Write-Host $selectedSolution.Solution -ForegroundColor $colorScheme.Default
                    
                    Write-Host ""
                    Write-Host "Press Enter to continue..." -ForegroundColor $colorScheme.Info
                    [void][Console]::ReadKey($true)
                }
            }
            3 { # Add Custom Solution
                Add-CustomSolution
            }
            4 { # Exit
                Clear-Host
                Write-Host "Thank you for using PowerShell Log Interpreter!" -ForegroundColor $colorScheme.Success
                return
            }
        }
    }
}

# Run the log interpreter
Start-LogInterpreter

