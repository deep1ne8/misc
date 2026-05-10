# Azure AD Migration Project Tracker
# This script creates and configures an Excel-based tracking system with PowerShell automation

# ===== PART 1: EXCEL TRACKER SETUP =====

function New-ExcelTemplate {
    param (
        [string]$OutputPath = "$env:USERPROFILE\Documents\AzureAD_Migration_Tracker.xlsx"
    )
    
    Write-Host "Creating Excel template for Azure AD Migration tracking..." -ForegroundColor Cyan
    
    # Verify Excel is installed
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
    }
    catch {
        Write-Error "Microsoft Excel is required for this solution. Please install Excel and try again."
        return
    }
    
    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Create main tracking worksheet
    $mainSheet = $workbook.Worksheets.Item(1)
    $mainSheet.Name = "Migration Tracking"
    
    # Set up column headers
    $headers = @(
        "ID", 
        "End User", 
        "Email", 
        "Phone", 
        "Workstation Name", 
        "Workstation Status", 
        "Scheduled Date", 
        "Scheduled Time", 
        "Data Migration Status", 
        "AD Disjoin Status", 
        "Azure AD Join Status", 
        "Profile Setup Status", 
        "Printer Installation Status", 
        "Special Notes", 
        "Last Updated", 
        "Updated By"
    )
    
    # Add headers
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $mainSheet.Cells.Item(1, $i + 1) = $headers[$i]
        $mainSheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $mainSheet.Cells.Item(1, $i + 1).Interior.ColorIndex = 15 # Light gray
    }
    
    # Format columns
    $mainSheet.Columns.Item(1).ColumnWidth = 5      # ID
    $mainSheet.Columns.Item(2).ColumnWidth = 20     # End User
    $mainSheet.Columns.Item(3).ColumnWidth = 25     # Email
    $mainSheet.Columns.Item(4).ColumnWidth = 15     # Phone
    $mainSheet.Columns.Item(5).ColumnWidth = 20     # Workstation Name
    $mainSheet.Columns.Item(6).ColumnWidth = 15     # Workstation Status
    $mainSheet.Columns.Item(7).ColumnWidth = 12     # Scheduled Date
    $mainSheet.Columns.Item(8).ColumnWidth = 12     # Scheduled Time
    $mainSheet.Columns.Item(9).ColumnWidth = 15     # Data Migration Status
    $mainSheet.Columns.Item(10).ColumnWidth = 15    # AD Disjoin Status
    $mainSheet.Columns.Item(11).ColumnWidth = 15    # Azure AD Join Status
    $mainSheet.Columns.Item(12).ColumnWidth = 15    # Profile Setup Status
    $mainSheet.Columns.Item(13).ColumnWidth = 15    # Printer Installation Status
    $mainSheet.Columns.Item(14).ColumnWidth = 30    # Special Notes
    $mainSheet.Columns.Item(15).ColumnWidth = 15    # Last Updated
    $mainSheet.Columns.Item(16).ColumnWidth = 15    # Updated By
    
    # Create data validation lists
    # Workstation Status
    $workstationStatusList = @("Not Started", "Pending", "In Progress", "Complete", "Issues Encountered")
    $statusRange = $mainSheet.Range("F2:F1000")
    Add-DataValidationList -Range $statusRange -ListValues $workstationStatusList
    
    # Task Status columns (columns I through M)
    $taskStatusList = @("Not Started", "In Progress", "Complete", "Error", "N/A")
    for ($col = 9; $col -le 13; $col++) {
        $statusRange = $mainSheet.Range(($mainSheet.Cells.Item(2, $col)), ($mainSheet.Cells.Item(1000, $col)))
        Add-DataValidationList -Range $statusRange -ListValues $taskStatusList
    }
    
    # Create History worksheet
    $historySheet = $workbook.Worksheets.Add()
    $historySheet.Name = "Migration History"
    
    # Set up history headers
    $historyHeaders = @(
        "Timestamp", 
        "ID", 
        "Workstation", 
        "End User", 
        "Action", 
        "Status", 
        "Comments", 
        "Performed By"
    )
    
    # Add history headers
    for ($i = 0; $i -lt $historyHeaders.Count; $i++) {
        $historySheet.Cells.Item(1, $i + 1) = $historyHeaders[$i]
        $historySheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $historySheet.Cells.Item(1, $i + 1).Interior.ColorIndex = 15 # Light gray
    }
    
    # Format history columns
    $historySheet.Columns.Item(1).ColumnWidth = 20  # Timestamp
    $historySheet.Columns.Item(2).ColumnWidth = 5   # ID
    $historySheet.Columns.Item(3).ColumnWidth = 20  # Workstation
    $historySheet.Columns.Item(4).ColumnWidth = 20  # End User
    $historySheet.Columns.Item(5).ColumnWidth = 25  # Action
    $historySheet.Columns.Item(6).ColumnWidth = 15  # Status
    $historySheet.Columns.Item(7).ColumnWidth = 40  # Comments
    $historySheet.Columns.Item(8).ColumnWidth = 15  # Performed By
    
    # Create Reports worksheet
    $reportsSheet = $workbook.Worksheets.Add()
    $reportsSheet.Name = "Reports"
    
    # Add report buttons and cells
    $reportsSheet.Cells.Item(1, 1) = "Azure AD Migration Project Reports"
    $reportsSheet.Cells.Item(1, 1).Font.Bold = $true
    $reportsSheet.Cells.Item(1, 1).Font.Size = 14
    
    $reportsSheet.Cells.Item(3, 1) = "Click buttons to generate reports:"
    $reportsSheet.Cells.Item(3, 1).Font.Bold = $true
    
    # Create Dashboard sheet
    $dashboardSheet = $workbook.Worksheets.Add()
    $dashboardSheet.Name = "Dashboard"
    
    # Set dashboard as first sheet
    $dashboardSheet.Move($workbook.Worksheets.Item(1))
    
    # Add dashboard elements
    $dashboardSheet.Cells.Item(1, 1) = "AZURE AD MIGRATION PROJECT DASHBOARD"
    $dashboardSheet.Cells.Item(1, 1).Font.Bold = $true
    $dashboardSheet.Cells.Item(1, 1).Font.Size = 16
    
    $dashboardSheet.Cells.Item(3, 1) = "Project Summary:"
    $dashboardSheet.Cells.Item(3, 1).Font.Bold = $true
    
    $dashboardSheet.Cells.Item(4, 1) = "Total Workstations:"
    $dashboardSheet.Cells.Item(4, 2) = "=COUNTA(Migration Tracking!E2:E1000)"
    
    $dashboardSheet.Cells.Item(5, 1) = "Total End Users:"
    $dashboardSheet.Cells.Item(5, 2) = "=COUNTA(Migration Tracking!B2:B1000)"
    
    $dashboardSheet.Cells.Item(6, 1) = "Migrations Complete:"
    $dashboardSheet.Cells.Item(6, 2) = "=COUNTIFS(Migration Tracking!F2:F1000,""Complete"")"
    
    $dashboardSheet.Cells.Item(7, 1) = "Migrations In Progress:"
    $dashboardSheet.Cells.Item(7, 2) = "=COUNTIFS(Migration Tracking!F2:F1000,""In Progress"")"
    
    $dashboardSheet.Cells.Item(8, 1) = "Migrations Pending:"
    $dashboardSheet.Cells.Item(8, 2) = "=COUNTIFS(Migration Tracking!F2:F1000,""Pending"")"
    
    $dashboardSheet.Cells.Item(9, 1) = "Issues Encountered:"
    $dashboardSheet.Cells.Item(9, 2) = "=COUNTIFS(Migration Tracking!F2:F1000,""Issues Encountered"")"
    
    # Format the summary section
    $summaryRange = $dashboardSheet.Range("A3:B9")
    $summaryRange.Borders.LineStyle = 1
    $dashboardSheet.Range("A3:A9").Font.Bold = $true
    
    # Create scheduled migrations section
    $dashboardSheet.Cells.Item(11, 1) = "Upcoming Scheduled Migrations:"
    $dashboardSheet.Cells.Item(11, 1).Font.Bold = $true
    
    $dashboardSheet.Cells.Item(12, 1) = "End User"
    $dashboardSheet.Cells.Item(12, 2) = "Workstation"
    $dashboardSheet.Cells.Item(12, 3) = "Scheduled Date"
    $dashboardSheet.Cells.Item(12, 4) = "Scheduled Time"
    
    $dashboardSheet.Range("A12:D12").Font.Bold = $true
    $dashboardSheet.Range("A12:D12").Interior.ColorIndex = 15
    
    # Formula for upcoming migrations (today's date or future)
    $today = Get-Date -Format "MM/dd/yyyy"
    $dashboardSheet.Range("A13:D17").FormulaArray = "=IFERROR(INDEX('Migration Tracking'!B:E,SMALL(IF('Migration Tracking'!G:G>=$today,ROW('Migration Tracking'!G:G)),ROW(1:5)),COLUMN()),"""")"
    
    # Format columns in dashboard
    $dashboardSheet.Columns.Item(1).ColumnWidth = 25
    $dashboardSheet.Columns.Item(2).ColumnWidth = 20
    $dashboardSheet.Columns.Item(3).ColumnWidth = 15
    $dashboardSheet.Columns.Item(4).ColumnWidth = 15
    
    # Add sample data for the first 3 users and 4 workstations
    $sampleUsers = @(
        @("John Smith", "john.smith@company.com", "555-123-4567", "WS-YELLOW-01"),
        @("Jane Doe", "jane.doe@company.com", "555-234-5678", "WS-YELLOW-02"),
        @("Alex Johnson", "alex.johnson@company.com", "555-345-6789", "WS-YELLOW-03", "WS-YELLOW-04")
    )
    
    $row = 2
    for ($i = 0; $i -lt $sampleUsers.Count; $i++) {
        $user = $sampleUsers[$i]
        
        if ($i -lt 2) {  # First two users have one workstation each
            $mainSheet.Cells.Item($row, 1) = $row - 1  # ID
            $mainSheet.Cells.Item($row, 2) = $user[0]  # User name
            $mainSheet.Cells.Item($row, 3) = $user[1]  # Email
            $mainSheet.Cells.Item($row, 4) = $user[2]  # Phone
            $mainSheet.Cells.Item($row, 5) = $user[3]  # Workstation
            $mainSheet.Cells.Item($row, 6) = "Pending" # Status
            $mainSheet.Cells.Item($row, 7) = (Get-Date).AddDays($i + 1).ToString("MM/dd/yyyy") # Schedule date
            $mainSheet.Cells.Item($row, 8) = "10:00 AM" # Schedule time
            $row++
        }
        else {  # Last user has two workstations
            for ($j = 3; $j -le 4; $j++) {
                $mainSheet.Cells.Item($row, 1) = $row - 1  # ID
                $mainSheet.Cells.Item($row, 2) = $user[0]  # User name
                $mainSheet.Cells.Item($row, 3) = $user[1]  # Email
                $mainSheet.Cells.Item($row, 4) = $user[2]  # Phone
                $mainSheet.Cells.Item($row, 5) = $user[$j]  # Workstation
                $mainSheet.Cells.Item($row, 6) = "Pending" # Status
                $mainSheet.Cells.Item($row, 7) = (Get-Date).AddDays($i + $j - 2).ToString("MM/dd/yyyy") # Schedule date
                $mainSheet.Cells.Item($row, 8) = "2:00 PM" # Schedule time
                $row++
            }
        }
    }
    
    # Add conditional formatting for status columns
    # Workstation Status (column F)
    Add-ConditionalFormatting -Worksheet $mainSheet -ColumnIndex 6 -StatusValues $workstationStatusList
    
    # Task Status columns (columns I through M)
    for ($col = 9; $col -le 13; $col++) {
        Add-ConditionalFormatting -Worksheet $mainSheet -ColumnIndex $col -StatusValues $taskStatusList
    }
    
    # Save the workbook
    try {
        $workbook.SaveAs($OutputPath)
        Write-Host "Excel template has been created successfully at: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to save the Excel file: $_"
    }
    finally {
        # Close Excel
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mainSheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($historySheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($reportsSheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dashboardSheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
    return $OutputPath
}

# Helper function to add data validation lists
function Add-DataValidationList {
    param (
        $Range,
        [string[]]$ListValues
    )
    
    $listString = $ListValues -join ","
    $Range.Validation.Delete()
    $Range.Validation.Add(3, 1, 1, $listString) | Out-Null
    $Range.Validation.IgnoreBlank = $true
    $Range.Validation.InCellDropdown = $true
}

# Helper function to add conditional formatting
function Add-ConditionalFormatting {
    param (
        $Worksheet,
        [int]$ColumnIndex,
        [string[]]$StatusValues
    )
    
    $range = $Worksheet.Range($Worksheet.Cells.Item(2, $ColumnIndex), $Worksheet.Cells.Item(1000, $ColumnIndex))
    
    # Set conditional formatting based on status values
    # Not Started - Light Gray
    $condition1 = $range.FormatConditions.Add(1, 3, "=""Not Started""")
    $condition1.Interior.ColorIndex = 15  # Light Gray
    
    # Pending - Yellow
    $condition2 = $range.FormatConditions.Add(1, 3, "=""Pending""")
    $condition2.Interior.ColorIndex = 6   # Yellow
    
    # In Progress - Light Blue
    $condition3 = $range.FormatConditions.Add(1, 3, "=""In Progress""")
    $condition3.Interior.ColorIndex = 8   # Light Blue
    
    # Complete - Green
    $condition4 = $range.FormatConditions.Add(1, 3, "=""Complete""")
    $condition4.Interior.ColorIndex = 4   # Green
    
    # Issues/Error - Red
    $condition5 = $range.FormatConditions.Add(1, 3, "=""Issues Encountered""")
    $condition5.Interior.ColorIndex = 3   # Red
    
    if ($StatusValues -contains "Error") {
        $condition6 = $range.FormatConditions.Add(1, 3, "=""Error""")
        $condition6.Interior.ColorIndex = 3   # Red
    }
    
    # N/A - Dark Gray
    if ($StatusValues -contains "N/A") {
        $condition7 = $range.FormatConditions.Add(1, 3, "=""N/A""")
        $condition7.Interior.ColorIndex = 16  # Dark Gray
    }
}

# ===== PART 2: POWERSHELL FUNCTIONS FOR TRACKER AUTOMATION =====

function Add-MigrationEntry {
    param (
        [string]$ExcelPath = "$env:USERPROFILE\Documents\AzureAD_Migration_Tracker.xlsx",
        [string]$EndUser,
        [string]$Email,
        [string]$Phone,
        [string]$WorkstationName,
        [string]$ScheduledDate,
        [string]$ScheduledTime = "10:00 AM",
        [string]$Notes = ""
    )
    
    Write-Host "Adding new migration entry..." -ForegroundColor Cyan
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($ExcelPath)
        $mainSheet = $workbook.Worksheets.Item("Migration Tracking")
        
        # Find the next available row
        $row = 2
        while ($null -ne $mainSheet.Cells.Item($row, 1).Value()) {
            $row++
        }
        
        # Add the new entry
        $mainSheet.Cells.Item($row, 1) = $row - 1 # ID
        $mainSheet.Cells.Item($row, 2) = $EndUser
        $mainSheet.Cells.Item($row, 3) = $Email
        $mainSheet.Cells.Item($row, 4) = $Phone
        $mainSheet.Cells.Item($row, 5) = $WorkstationName
        $mainSheet.Cells.Item($row, 6) = "Pending" # Default status
        $mainSheet.Cells.Item($row, 7) = $ScheduledDate
        $mainSheet.Cells.Item($row, 8) = $ScheduledTime
        $mainSheet.Cells.Item($row, 9) = "Not Started" # Data Migration Status
        $mainSheet.Cells.Item($row, 10) = "Not Started" # AD Disjoin Status
        $mainSheet.Cells.Item($row, 11) = "Not Started" # Azure AD Join Status
        $mainSheet.Cells.Item($row, 12) = "Not Started" # Profile Setup Status
        $mainSheet.Cells.Item($row, 13) = "Not Started" # Printer Installation Status
        $mainSheet.Cells.Item($row, 14) = $Notes
        $mainSheet.Cells.Item($row, 15) = Get-Date -Format "MM/dd/yyyy HH:mm"
        $mainSheet.Cells.Item($row, 16) = $env:USERNAME
        
        # Add entry to history
        $historySheet = $workbook.Worksheets.Item("Migration History")
        $historyRow = 2
        while ($null -ne $historySheet.Cells.Item($historyRow, 1).Value()) {
            $historyRow++
        }
        
        $historySheet.Cells.Item($historyRow, 1) = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
        $historySheet.Cells.Item($historyRow, 2) = $row - 1 # ID
        $historySheet.Cells.Item($historyRow, 3) = $WorkstationName
        $historySheet.Cells.Item($historyRow, 4) = $EndUser
        $historySheet.Cells.Item($historyRow, 5) = "Migration Entry Created"
        $historySheet.Cells.Item($historyRow, 6) = "Pending"
        $historySheet.Cells.Item($historyRow, 7) = "New migration entry added"
        $historySheet.Cells.Item($historyRow, 8) = $env:USERNAME
        
        # Save the workbook
        $workbook.Save()
        Write-Host "New migration entry added successfully with ID $($row - 1)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to add migration entry: $_"
    }
    finally {
        # Close Excel
        if ($null -ne $workbook) {
            $workbook.Close($true)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Update-MigrationStatus {
    param (
        [string]$ExcelPath = "$env:USERPROFILE\Documents\AzureAD_Migration_Tracker.xlsx",
        [int]$ID,
        [ValidateSet("Not Started", "Pending", "In Progress", "Complete", "Issues Encountered")]
        [string]$WorkstationStatus,
        [ValidateSet("Not Started", "In Progress", "Complete", "Error", "N/A")]
        [string]$DataMigrationStatus,
        [ValidateSet("Not Started", "In Progress", "Complete", "Error", "N/A")]
        [string]$ADDisjoinStatus,
        [ValidateSet("Not Started", "In Progress", "Complete", "Error", "N/A")]
        [string]$AzureADJoinStatus,
        [ValidateSet("Not Started", "In Progress", "Complete", "Error", "N/A")]
        [string]$ProfileSetupStatus,
        [ValidateSet("Not Started", "In Progress", "Complete", "Error", "N/A")]
        [string]$PrinterInstallStatus,
        [string]$Notes = "",
        [switch]$AddToHistory
    )
    
    Write-Host "Updating migration status for ID $ID..." -ForegroundColor Cyan
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($ExcelPath)
        $mainSheet = $workbook.Worksheets.Item("Migration Tracking")
        
        # Find the row with the matching ID
        $found = $false
        $row = 2
        while ($null -ne $mainSheet.Cells.Item($row, 1).Value()) {
            if ($mainSheet.Cells.Item($row, 1).Value() -eq $ID) {
                $found = $true
                break
            }
            $row++
        }
        
        if (-not $found) {
            Write-Error "No migration entry found with ID $ID"
            return
        }
        
        # Get current values for history
        $currentWorkstation = $mainSheet.Cells.Item($row, 5).Value()
        $currentEndUser = $mainSheet.Cells.Item($row, 2).Value()
        
        # Update status values if provided
        if ($WorkstationStatus) {
            $mainSheet.Cells.Item($row, 6) = $WorkstationStatus
        }
        if ($DataMigrationStatus) {
            $mainSheet.Cells.Item($row, 9) = $DataMigrationStatus
        }
        if ($ADDisjoinStatus) {
            $mainSheet.Cells.Item($row, 10) = $ADDisjoinStatus
        }
        if ($AzureADJoinStatus) {
            $mainSheet.Cells.Item($row, 11) = $AzureADJoinStatus
        }
        if ($ProfileSetupStatus) {
            $mainSheet.Cells.Item($row, 12) = $ProfileSetupStatus
        }
        if ($PrinterInstallStatus) {
            $mainSheet.Cells.Item($row, 13) = $PrinterInstallStatus
        }
        
        # Update notes if provided
        if ($Notes) {
            $mainSheet.Cells.Item($row, 14) = $Notes
        }
        
        # Update timestamp and user
        $mainSheet.Cells.Item($row, 15) = Get-Date -Format "MM/dd/yyyy HH:mm"
        $mainSheet.Cells.Item($row, 16) = $env:USERNAME
        
        # Add entry to history if requested
        if ($AddToHistory) {
            $historySheet = $workbook.Worksheets.Item("Migration History")
            $historyRow = 2
            while ($null -ne $historySheet.Cells.Item($historyRow, 1).Value()) {
                $historyRow++
            }
            
            $changes = @()
            if ($WorkstationStatus) { $changes += "Workstation Status: $WorkstationStatus" }
            if ($DataMigrationStatus) { $changes += "Data Migration: $DataMigrationStatus" }
            if ($ADDisjoinStatus) { $changes += "AD Disjoin: $ADDisjoinStatus" }
            if ($AzureADJoinStatus) { $changes += "Azure AD Join: $AzureADJoinStatus" }
            if ($ProfileSetupStatus) { $changes += "Profile Setup: $ProfileSetupStatus" }
            if ($PrinterInstallStatus) { $changes += "Printer Install: $PrinterInstallStatus" }
            
            $historySheet.Cells.Item($historyRow, 1) = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
            $historySheet.Cells.Item($historyRow, 2) = $ID
            $historySheet.Cells.Item($historyRow, 3) = $currentWorkstation
            $historySheet.Cells.Item($historyRow, 4) = $currentEndUser
            $historySheet.Cells.Item($historyRow, 5) = "Status Update"
            $historySheet.Cells.Item($historyRow, 6) = $WorkstationStatus
            $historySheet.Cells.Item($historyRow, 7) = "Updated: " + ($changes -join ", ") + $(if ($Notes) { "; Notes: $Notes" })
            $historySheet.Cells.Item($historyRow, 8) = $env:USERNAME
        }
        
        # Save the workbook
        $workbook.Save()
        Write-Host "Migration status updated successfully for ID $ID" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to update migration status: $_"
    }
    finally {
        # Close Excel
        if ($null -ne $workbook) {
            $workbook.Close($true)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Export-MigrationReport {
    param (
        [string]$ExcelPath = "$env:USERPROFILE\Documents\AzureAD_Migration_Tracker.xlsx",
        [string]$OutputPath = "$env:USERPROFILE\Documents\AzureAD_Migration_Report_$(Get-Date -Format 'yyyyMMdd').html",
        [switch]$IncludeHistory
    )
    
    Write-Host "Generating migration report..." -ForegroundColor Cyan
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($ExcelPath)
        $mainSheet = $workbook.Worksheets.Item("Migration Tracking")
        
        # Get migration data
        $migrations = @()
        $row = 2
        while ($null -ne $mainSheet.Cells.Item($row, 1).Value()) {
            $migration = [PSCustomObject]@{
                ID = $mainSheet.Cells.Item($row, 1).Value()
                EndUser = $mainSheet.Cells.Item($row, 2).Value()
                Email = $mainSheet.Cells.Item($row, 3).Value()
                WorkstationName = $mainSheet.Cells.Item($row, 5).Value()
                WorkstationStatus = $mainSheet.Cells.Item($row, 6).Value()
                ScheduledDate = $mainSheet.Cells.Item($row, 7).Value()
                ScheduledTime = $mainSheet.Cells.Item($row, 8).Value()
                DataMigrationStatus = $mainSheet.Cells.Item($row, 9).Value()
                ADDisjoinStatus = $mainSheet.Cells.Item($row, 10).Value()
                AzureADJoinStatus = $mainSheet.Cells.Item($row, 11).Value()
                ProfileSetupStatus = $mainSheet.Cells.Item($row, 12).Value()
                PrinterInstallStatus = $mainSheet.Cells.Item($row, 13).Value()
                Notes = $mainSheet.Cells.Item($row, 14).Value()
                LastUpdated = $mainSheet.Cells.Item($row, 15).Value()
                UpdatedBy = $mainSheet.Cells.Item($row, 16).Value()
            }
            $migrations += $migration
            $row++
        }
        
        # Get history data if requested
        $history = @()
        if ($IncludeHistory) {
            $historySheet = $workbook.Worksheets.Item("Migration History")
            $row = 2
            while ($null -ne $historySheet.Cells.Item($row, 1).Value()) {
                $entry = [PSCustomObject]@{
                    Timestamp = $historySheet.Cells.Item($row, 1).Value()
                    ID = $historySheet.Cells.Item($row, 2).Value()
                    Workstation = $historySheet.Cells.Item($row, 3).Value()
                    EndUser = $historySheet.Cells.Item($row, 4).Value()
                    Action = $historySheet.Cells.Item($row, 5).Value()
                    Status = $historySheet.Cells.Item($row, 6).Value()
                    Comments = $historySheet.Cells.Item($row, 7).Value()
                    PerformedBy = $historySheet.Cells.Item($row, 8).Value()
                }
                $history += $entry
                $row++
            }
        }
        
        # Generate HTML report
        $reportTitle = "Azure AD Migration Project Report - $(Get-Date -Format 'yyyy-MM-dd')"
        $css = @"
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            h1, h2 { color: #0078D4; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
            th { background-color: #0078D4; color: white; text-align: left; padding: 8px; }
            td { border: 1px solid #ddd; padding: 8px; }
            tr:nth-child(even) { background-color: #f2f2f2; }
            .status-complete { background-color: #DFF6DD; }
            .status-inprogress { background-color: #EFF6FC; }
            .status-pending { background-color: #FFF4CE; }
            .status-error { background-color: #FDE7E9; }
            .status-notstarted { background-color: #f0f0f0; }
            .summary { display: flex; flex-wrap: wrap; margin-bottom: 20px; }
            .summary-item { 
                margin: 10px; padding: 15px; border-radius: 5px; 
                box-shadow: 0 2px 5px rgba(0,0,0,0.1); min-width: 150px; text-align: center; 
            }
            .summary-count { font-size: 24px; font-weight: bold; margin: 10px 0; }
            .print-button { 
                background-color: #0078D4; color: white; border: none; 
                padding: 10px 15px; cursor: pointer; margin-bottom: 20px;
                border-radius: 4px;
            }
        </style>
"@
        
        # Create summary statistics
        $totalWorkstations = $migrations.Count
        $complete = ($migrations | Where-Object { $_.WorkstationStatus -eq "Complete" }).Count
        $inProgress = ($migrations | Where-Object { $_.WorkstationStatus -eq "In Progress" }).Count
        $pending = ($migrations | Where-Object { $_.WorkstationStatus -eq "Pending" }).Count
        $issues = ($migrations | Where-Object { $_.WorkstationStatus -eq "Issues Encountered" }).Count
        
        $summaryHtml = @"
        <div class="summary">
            <div class="summary-item">
                <div>Total Workstations</div>
                <div class="summary-count">$totalWorkstations</div>
            </div>
            <div class="summary-item status-complete">
                <div>Complete</div>
                <div class="summary-count">$complete</div>
            </div>
            <div class="summary-item status-inprogress">
                <div>In Progress</div>
                <div class="summary-count">$inProgress</div>
            </div>
            <div class="summary-item status-pending">
                <div>Pending</div>
                <div class="summary-count">$pending</div>
            </div>
            <div class="summary-item status-error">
                <div>Issues</div>
                <div class="summary-count">$issues</div>
            </div>
        </div>
"@
        
        # Create migration table
        $migrationsHtml = @"
        <h2>Migration Status</h2>
        <table>
            <tr>
                <th>ID</th>
                <th>End User</th>
                <th>Workstation</th>
                <th>Status</th>
                <th>Scheduled</th>
                <th>Data Migration</th>
                <th>AD Disjoin</th>
                <th>Azure AD Join</th>
                <th>Profile Setup</th>
                <th>Printer Install</th>
                <th>Last Updated</th>
            </tr>
"@
        
        foreach ($migration in $migrations) {
            $statusClass = switch ($migration.WorkstationStatus) {
                "Complete" { "status-complete" }
                "In Progress" { "status-inprogress" }
                "Pending" { "status-pending" }
                "Issues Encountered" { "status-error" }
                "Not Started" { "status-notstarted" }
                default { "" }
            }
            
            $migrationsHtml += @"
            <tr class="$statusClass">
                <td>$($migration.ID)</td>
                <td>$($migration.EndUser)</td>
                <td>$($migration.WorkstationName)</td>
                <td>$($migration.WorkstationStatus)</td>
                <td>$($migration.ScheduledDate) $($migration.ScheduledTime)</td>
                <td>$($migration.DataMigrationStatus)</td>
                <td>$($migration.ADDisjoinStatus)</td>
                <td>$($migration.AzureADJoinStatus)</td>
                <td>$($migration.ProfileSetupStatus)</td>
                <td>$($migration.PrinterInstallStatus)</td>
                <td>$($migration.LastUpdated)</td>
            </tr>
"@
        }
        $migrationsHtml += "</table>"
        
        # Create history table if requested
        $historyHtml = ""
        if ($IncludeHistory -and $history.Count -gt 0) {
            $historyHtml = @"
            <h2>Migration History</h2>
            <table>
                <tr>
                    <th>Timestamp</th>
                    <th>ID</th>
                    <th>Workstation</th>
                    <th>End User</th>
                    <th>Action</th>
                    <th>Status</th>
                    <th>Comments</th>
                    <th>Performed By</th>
                </tr>
"@
            
            foreach ($entry in $history) {
                $historyHtml += @"
                <tr>
                    <td>$($entry.Timestamp)</td>
                    <td>$($entry.ID)</td>
                    <td>$($entry.Workstation)</td>
                    <td>$($entry.EndUser)</td>
                    <td>$($entry.Action)</td>
                    <td>$($entry.Status)</td>
                    <td>$($entry.Comments)</td>
                    <td>$($entry.PerformedBy)</td>
                </tr>
"@
            }
            $historyHtml += "</table>"
        }
        
        # Combine everything into a complete HTML report
        $html = @"
        <!DOCTYPE html>
        <!--
            Please report any bugs such as null pointer references, unhandled exceptions, and more to 
            https://github.com/deep1ne8/misc/issues
        -->
        <html>
        <head>
            <title>$reportTitle</title>
            $css
            <script>
                function printReport() {
                    window.print();
                }
            </script>
        </head>
        <body>
            <h1>$reportTitle</h1>
            <button class="print-button" onclick="printReport()">Print Report</button>
            $summaryHtml
            $migrationsHtml
            $historyHtml
        </body>
        </html>
"@
        
        # Save the HTML report
        $html | Out-File -FilePath $OutputPath -Encoding utf8
        
        Write-Host "Migration report generated successfully at: $OutputPath" -ForegroundColor Green
        
        # Open the report in the default browser
        Start-Process $OutputPath
    }
    catch {
        Write-Error "Failed to generate migration report: $_"
    }
    finally {
        # Close Excel
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ===== PART 3: MAIN EXECUTION =====

# Create the initial Excel template
$excelPath = New-ExcelTemplate

Write-Host "Azure AD Migration Project Tracker has been set up successfully!" -ForegroundColor Green
Write-Host "Excel file location: $excelPath" -ForegroundColor Yellow
Write-Host ""
Write-Host "Available Functions:" -ForegroundColor Cyan
Write-Host "- Add-MigrationEntry: Add a new migration task" -ForegroundColor White
Write-Host "- Update-MigrationStatus: Update the status of a migration task" -ForegroundColor White
Write-Host "- Export-MigrationReport: Generate an HTML report of migration status" -ForegroundColor White
Write-Host ""
Write-Host "Example usage:" -ForegroundColor Cyan
Write-Host "Add-MigrationEntry -EndUser 'John Smith' -Email 'john.smith@company.com' -Phone '555-123-4567' -WorkstationName 'WS-YELLOW-01' -ScheduledDate '03/01/2025'" -ForegroundColor White
Write-Host "Update-MigrationStatus -ID 1 -WorkstationStatus 'In Progress' -DataMigrationStatus 'Complete' -Notes 'User data backed up successfully'" -ForegroundColor White
Write-Host "Export-MigrationReport -IncludeHistory" -ForegroundColor White
