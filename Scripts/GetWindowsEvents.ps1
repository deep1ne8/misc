#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Windows Event Log Query Tool with Color-coded Output
.DESCRIPTION
    Provides menu-driven interface to query Application and System event logs
    with filtering options and detailed event inspection capabilities
.AUTHOR
    Generated PowerShell Script
#>

# Color scheme for event levels
$ColorMap = @{
    'Critical' = 'Red'
    'Error' = 'Red'
    'Warning' = 'Yellow'  
    'Information' = 'Green'
    'Verbose' = 'Cyan'
    'LogAlways' = 'White'
}

function Show-MainMenu {
    Clear-Host
    Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "    Windows Event Log Query Tool" -ForegroundColor White
    Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select Event Log:" -ForegroundColor Yellow
    Write-Host "[1] Application Log" -ForegroundColor White
    Write-Host "[2] System Log" -ForegroundColor White
    Write-Host "[0] Exit" -ForegroundColor Red
    Write-Host ""
}

function Show-FilterMenu {
    param([string]$LogName)
    
    Clear-Host
    Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "    $LogName Event Log Filters" -ForegroundColor White
    Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select Filter Option:" -ForegroundColor Yellow
    Write-Host "[1] All Events" -ForegroundColor White
    Write-Host "[2] Critical Only" -ForegroundColor Red
    Write-Host "[3] Error Only" -ForegroundColor Red
    Write-Host "[4] Critical + Error" -ForegroundColor Red
    Write-Host "[5] Warning + Error + Critical" -ForegroundColor Yellow
    Write-Host "[6] Custom Date Range" -ForegroundColor Magenta
    Write-Host "[B] Back to Main Menu" -ForegroundColor Green
    Write-Host "[0] Exit" -ForegroundColor Red
    Write-Host ""
}

function Show-EventDetails {
    param([object]$Event)
    
    $LevelName = switch ($Event.Level) {
        1 { "Critical" }
        2 { "Error" }
        3 { "Warning" }
        4 { "Information" }
        5 { "Verbose" }
        default { "Unknown" }
    }
    
    $Color = $ColorMap[$LevelName]
    if (-not $Color) { $Color = 'White' }
    
    Write-Host "─────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "Event ID: " -NoNewline -ForegroundColor White
    Write-Host $Event.Id -ForegroundColor Cyan
    Write-Host "Level: " -NoNewline -ForegroundColor White
    Write-Host $LevelName -ForegroundColor $Color
    Write-Host "Time Created: " -NoNewline -ForegroundColor White
    Write-Host $Event.TimeCreated -ForegroundColor White
    Write-Host "Source: " -NoNewline -ForegroundColor White
    Write-Host $Event.ProviderName -ForegroundColor Yellow
    Write-Host "Task Category: " -NoNewline -ForegroundColor White
    Write-Host $Event.TaskDisplayName -ForegroundColor White
    Write-Host ""
    Write-Host "Message:" -ForegroundColor White
    Write-Host $Event.Message -ForegroundColor Gray -Wrap
    Write-Host "─────────────────────────────────────────" -ForegroundColor Gray
}

function Get-FilteredEvents {
    param(
        [string]$LogName,
        [int]$FilterOption,
        [datetime]$StartTime = $null,
        [datetime]$EndTime = $null,
        [int]$MaxEvents = 50
    )
    
    try {
        $FilterHashtable = @{
            LogName = $LogName
            MaxEvents = $MaxEvents
        }
        
        # Add time filter if specified
        if ($StartTime) { $FilterHashtable.StartTime = $StartTime }
        if ($EndTime) { $FilterHashtable.EndTime = $EndTime }
        
        # Add level filter based on selection
        switch ($FilterOption) {
            2 { $FilterHashtable.Level = 1 }  # Critical
            3 { $FilterHashtable.Level = 2 }  # Error
            4 { $FilterHashtable.Level = 1,2 }  # Critical + Error
            5 { $FilterHashtable.Level = 1,2,3 }  # Warning + Error + Critical
        }
        
        Write-Host "Querying events..." -ForegroundColor Yellow
        $Events = Get-WinEvent -FilterHashtable $FilterHashtable -ErrorAction Stop
        
        if ($Events.Count -eq 0) {
            Write-Host "No events found matching the criteria." -ForegroundColor Yellow
            return $null
        }
        
        return $Events
    }
    catch {
        Write-Host "Error querying events: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Show-EventList {
    param([array]$Events)
    
    if (-not $Events) { return }
    
    Write-Host ""
    Write-Host "Found $($Events.Count) events:" -ForegroundColor Green
    Write-Host ""
    
    for ($i = 0; $i -lt $Events.Count; $i++) {
        $Event = $Events[$i]
        $LevelName = switch ($Event.Level) {
            1 { "Critical" }
            2 { "Error" }
            3 { "Warning" }
            4 { "Information" }
            5 { "Verbose" }
            default { "Unknown" }
        }
        
        $Color = $ColorMap[$LevelName]
        if (-not $Color) { $Color = 'White' }
        
        $IndexDisplay = "[$($i + 1)]".PadRight(6)
        $TimeDisplay = $Event.TimeCreated.ToString("MM/dd HH:mm:ss")
        $IdDisplay = "ID:$($Event.Id)".PadRight(10)
        $LevelDisplay = $LevelName.PadRight(11)
        $SourceDisplay = $Event.ProviderName
        
        Write-Host $IndexDisplay -NoNewline -ForegroundColor White
        Write-Host $TimeDisplay -NoNewline -ForegroundColor Gray
        Write-Host " $IdDisplay" -NoNewline -ForegroundColor Cyan
        Write-Host $LevelDisplay -NoNewline -ForegroundColor $Color
        Write-Host " $SourceDisplay" -ForegroundColor Yellow
    }
}

function Show-EventMenu {
    param([array]$Events, [string]$LogName)
    
    do {
        Write-Host ""
        Write-Host "─────────────────────────────────────────" -ForegroundColor Gray
        Write-Host "Enter event number for details, or:" -ForegroundColor Yellow
        Write-Host "[R] Refresh query" -ForegroundColor Green
        Write-Host "[B] Back to filter menu" -ForegroundColor Green  
        Write-Host "[M] Main menu" -ForegroundColor Green
        Write-Host "[0] Exit" -ForegroundColor Red
        
        $Choice = Read-Host "Choice"
        
        switch ($Choice.ToUpper()) {
            'R' { return 'REFRESH' }
            'B' { return 'BACK' }
            'M' { return 'MAIN' }
            '0' { return 'EXIT' }
            default {
                if ([int]::TryParse($Choice, [ref]$null) -and [int]$Choice -ge 1 -and [int]$Choice -le $Events.Count) {
                    Clear-Host
                    Write-Host "Event Details - $LogName Log" -ForegroundColor Cyan
                    Show-EventDetails -Event $Events[[int]$Choice - 1]
                    Write-Host ""
                    Write-Host "Press any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    Show-EventList -Events $Events
                }
                else {
                    Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                }
            }
        }
    } while ($true)
}

function Get-CustomDateRange {
    Write-Host ""
    Write-Host "Custom Date Range Filter" -ForegroundColor Cyan
    Write-Host "Format: MM/dd/yyyy or MM/dd/yyyy HH:mm" -ForegroundColor Gray
    Write-Host ""
    
    do {
        $StartInput = Read-Host "Start date/time (press Enter for 24 hours ago)"
        if ([string]::IsNullOrWhiteSpace($StartInput)) {
            $StartTime = (Get-Date).AddDays(-1)
            break
        }
        elseif ([datetime]::TryParse($StartInput, [ref]$StartTime)) {
            break
        }
        else {
            Write-Host "Invalid date format. Please try again." -ForegroundColor Red
        }
    } while ($true)
    
    do {
        $EndInput = Read-Host "End date/time (press Enter for now)"
        if ([string]::IsNullOrWhiteSpace($EndInput)) {
            $EndTime = Get-Date
            break
        }
        elseif ([datetime]::TryParse($EndInput, [ref]$EndTime)) {
            break
        }
        else {
            Write-Host "Invalid date format. Please try again." -ForegroundColor Red
        }
    } while ($true)
    
    return @{
        StartTime = $StartTime
        EndTime = $EndTime
    }
}

# Main script execution
try {
    # Check if running as administrator for better event access
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $IsAdmin) {
        Write-Host "Note: Running without administrator privileges may limit event access." -ForegroundColor Yellow
        Write-Host ""
    }
    
    do {
        Show-MainMenu
        $MainChoice = Read-Host "Select option"
        
        switch ($MainChoice) {
            '1' { $LogName = 'Application' }
            '2' { $LogName = 'System' }
            '0' { 
                Write-Host "Exiting..." -ForegroundColor Green
                exit 
            }
            default { 
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 1
                continue
            }
        }
        
        # Filter menu loop
        do {
            Show-FilterMenu -LogName $LogName
            $FilterChoice = Read-Host "Select filter option"
            
            switch ($FilterChoice.ToUpper()) {
                'B' { break }  # Back to main menu
                '0' { 
                    Write-Host "Exiting..." -ForegroundColor Green
                    exit 
                }
                { $_ -in '1','2','3','4','5','6' } {
                    $StartTime = $null
                    $EndTime = $null
                    
                    # Handle custom date range
                    if ($FilterChoice -eq '6') {
                        $DateRange = Get-CustomDateRange
                        $StartTime = $DateRange.StartTime
                        $EndTime = $DateRange.EndTime
                        $FilterOption = 1  # Show all events in date range
                    }
                    else {
                        $FilterOption = [int]$FilterChoice
                    }
                    
                    # Query events
                    $Events = Get-FilteredEvents -LogName $LogName -FilterOption $FilterOption -StartTime $StartTime -EndTime $EndTime
                    
                    if ($Events) {
                        do {
                            Clear-Host
                            Write-Host "Events from $LogName Log" -ForegroundColor Cyan
                            if ($StartTime) {
                                Write-Host "Date Range: $($StartTime.ToString('MM/dd/yyyy HH:mm')) to $($EndTime.ToString('MM/dd/yyyy HH:mm'))" -ForegroundColor Gray
                            }
                            Show-EventList -Events $Events
                            
                            $EventMenuResult = Show-EventMenu -Events $Events -LogName $LogName
                            
                            switch ($EventMenuResult) {
                                'REFRESH' { 
                                    $Events = Get-FilteredEvents -LogName $LogName -FilterOption $FilterOption -StartTime $StartTime -EndTime $EndTime
                                    if (-not $Events) { break }
                                }
                                'BACK' { break }
                                'MAIN' { 
                                    $FilterChoice = 'B'  # This will break both loops
                                    break 
                                }
                                'EXIT' { 
                                    Write-Host "Exiting..." -ForegroundColor Green
                                    exit 
                                }
                            }
                        } while ($EventMenuResult -notin @('BACK', 'MAIN'))
                    }
                    else {
                        Write-Host "Press any key to continue..." -ForegroundColor Gray
                        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    }
                }
                default { 
                    Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
        } while ($FilterChoice.ToUpper() -ne 'B')
        
    } while ($true)
}
catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
