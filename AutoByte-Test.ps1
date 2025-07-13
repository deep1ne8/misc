# AutoBytePro GUI Script - Enhanced Version with PSScriptAnalyzer Fixes
# Description: GUI to select and run PowerShell scripts from GitHub with real-time output
# Author: Enhanced for improved functionality and error handling

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Set application settings for better stability - MUST be called before creating any forms
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# Global variables
$script:currentProcess = $null
$script:tempFiles = @()

# Define scripts with improved structure
$GitHubScripts = @(
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1"; Description = "Disk Cleaner"; Category = "Maintenance" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1"; Description = "Enable Files On Demand"; Category = "Configuration" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadandInstallPackage.ps1"; Description = "Download & Install Package"; Category = "Installation" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckUserProfileIssue.ps1"; Description = "Check User Profile"; Category = "Diagnostics" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1"; Description = "Dell Bloatware Remover"; Category = "Maintenance" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1"; Description = "Reset & Install Windows Update"; Category = "Updates" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1"; Description = "Windows System Repair"; Category = "Repair" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Description = "Reset Windows Search DB"; Category = "Repair" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallMSProjects.ps1"; Description = "Install MS Projects"; Category = "Installation" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1"; Description = "Check Drive Space"; Category = "Diagnostics" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1"; Description = "Internet Speed Test"; Category = "Network" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1"; Description = "Internet Latency Test"; Category = "Network" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WorkPaperMonitorTroubleShooter.ps1"; Description = "Monitor Troubleshooter"; Category = "Hardware" }
)

# Function to safely append text to RichTextBox
$script:WriteOutputDelegate = {
    param([string]$Message, [string]$Color = "Black")
    
    if ($richTextBox.InvokeRequired) {
        $richTextBox.Invoke({ param($msg, $clr) Write-Output -Message $msg -Color $clr }, $Message, $Color)
        return
    }
    
    $richTextBox.SelectionStart = $richTextBox.TextLength
    $richTextBox.SelectionLength = 0
    $richTextBox.SelectionColor = [System.Drawing.Color]::FromName($Color)
    $richTextBox.AppendText("$Message`r`n")
    $richTextBox.SelectionColor = [System.Drawing.Color]::Black
    $richTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-Output {
    param([string]$Message, [string]$Color = "Black")
    
    if ($richTextBox.InvokeRequired) {
        $richTextBox.Invoke($script:WriteOutputDelegate, $Message, $Color)
        return
    }
    
    & $script:WriteOutputDelegate $Message $Color
}

# Function to cleanup temporary files
function Remove-TempFiles {
    foreach ($file in $script:tempFiles) {
        if (Test-Path $file) {
            try {
                Remove-Item $file -Force -ErrorAction SilentlyContinue
                Write-Output "Cleaned up: $file" "Gray"
            } catch {
                Write-Output "Warning: Could not remove $file" "Orange"
            }
        }
    }
    $script:tempFiles = @()
}

# Function to stop current process
function Stop-CurrentProcess {
    try {
        if ($script:currentProcess -and !$script:currentProcess.HasExited) {
            $script:currentProcess.Kill()
            Write-Output "Process terminated by user" "Red"
        }
    } catch {
        Write-Output "Error terminating process: $($_.Exception.Message)" "Red"
    } finally {
        # Always re-enable buttons
        if ($buttonPanel) { $buttonPanel.Enabled = $true }
        if ($stopButton) { $stopButton.Enabled = $false }
        Update-Status "Ready"
    }
}

# Function to update status
function Update-Status {
    param([string]$Message)
    try {
        if ($statusLabel -and $statusLabel.Owner) {
            if ($statusLabel.Owner.InvokeRequired) {
                $statusLabel.Owner.Invoke([Action]{
                    $statusLabel.Text = $Message
                })
            } else {
                $statusLabel.Text = $Message
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    } catch {
        # Silently handle any status update errors
    }
}

# Enhanced function to download, save, and execute script with better error handling
function Start-Script {
    param([string]$Url, [string]$Description)
    
    # Clear output and disable buttons during execution
    $richTextBox.Clear()
    $buttonPanel.Enabled = $false
    $stopButton.Enabled = $true
    Update-Status "Running: $Description"
    
    try {
        Write-Output "===========================================" "Blue"
        Write-Output "Starting: $Description" "Blue"
        Write-Output "===========================================" "Blue"
        Write-Output "Downloading script from: $Url" "Green"
        
        # Download with timeout and progress
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "AutoBytePro/1.0")
        $scriptContent = $webClient.DownloadString($Url)
        $webClient.Dispose()
        
        if ([string]::IsNullOrWhiteSpace($scriptContent)) {
            throw "Downloaded script is empty or invalid"
        }
        
        Write-Output "Script downloaded successfully ($(($scriptContent.Length / 1KB).ToString('F2')) KB)" "Green"
        
        # Create temporary file with unique name
        $tempFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), "AutoBytePro_$([guid]::NewGuid()).ps1")
        $script:tempFiles += $tempFile
        
        Set-Content -Path $tempFile -Value $scriptContent -Encoding UTF8 -Force
        Write-Output "Saved to temporary file: $tempFile" "Gray"
        Write-Output "Executing script..." "Green"
        Write-Output "===========================================" "Blue"
        
        # Configure process with better settings
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$tempFile`""
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = [System.Environment]::CurrentDirectory
        
        $script:currentProcess = New-Object System.Diagnostics.Process
        $script:currentProcess.StartInfo = $psi
        $script:currentProcess.EnableRaisingEvents = $true
        
        # Event handlers for output - FIXED: Use different parameter names
        $script:currentProcess.add_OutputDataReceived({
            param($processObject, $dataArgs)
            if ($dataArgs.Data) {
                $richTextBox.Invoke([Action]{
                    Write-Output $dataArgs.Data "Black"
                })
            }
        })
        
        $script:currentProcess.add_ErrorDataReceived({
            param($processObject, $errorArgs)
            if ($errorArgs.Data) {
                $richTextBox.Invoke([Action]{
                    Write-Output "ERROR: $($errorArgs.Data)" "Red"
                })
            }
        })
        
        $script:currentProcess.add_Exited({
            param($processObject, $exitArgs)
            try {
                $exitCode = $script:currentProcess.ExitCode
                $richTextBox.Invoke([Action]{
                    Write-Output "===========================================" "Blue"
                    if ($exitCode -eq 0) {
                        Write-Output "Script completed successfully (Exit Code: $exitCode)" "Green"
                        Update-Status "Completed: $Description"
                    } else {
                        Write-Output "Script finished with exit code: $exitCode" "Orange"
                        Update-Status "Failed: $Description"
                    }
                    Write-Output "===========================================" "Blue"
                    
                    # Re-enable buttons
                    $buttonPanel.Enabled = $true
                    $stopButton.Enabled = $false
                })
            } catch {
                # Handle any errors in the exit event
                Write-Output "Error in exit handler: $($_.Exception.Message)" "Red"
            }
        })
        
        # Start process
        $script:currentProcess.Start() | Out-Null
        $script:currentProcess.BeginOutputReadLine()
        $script:currentProcess.BeginErrorReadLine()
        
    } catch {
        Write-Output "CRITICAL ERROR: $($_.Exception.Message)" "Red"
        Write-Output "Stack Trace: $($_.Exception.StackTrace)" "Red"
        $buttonPanel.Enabled = $true
        $stopButton.Enabled = $false
        Update-Status "Error: $($_.Exception.Message)"
    }
}

# Create main form with improved design
$form = New-Object System.Windows.Forms.Form
$form.Text = 'AutoBytePro GUI v2.0'
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.Icon = [System.Drawing.SystemIcons]::Application

# Add form closing event
$form.add_FormClosing({
    Stop-CurrentProcess
    Remove-TempFiles
})

# Create menu strip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"
$clearOutputItem = New-Object System.Windows.Forms.ToolStripMenuItem
$clearOutputItem.Text = "Clear Output"
$clearOutputItem.ShortcutKeys = "Ctrl+L"
$clearOutputItem.add_Click({ $richTextBox.Clear(); Update-Status "Output cleared" })
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "E&xit"
$exitItem.ShortcutKeys = "Ctrl+Q"
$exitItem.add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($clearOutputItem)
$fileMenu.DropDownItems.Add($exitItem)
$menuStrip.Items.Add($fileMenu)

# Create toolbar
$toolStrip = New-Object System.Windows.Forms.ToolStrip
$toolStrip.GripStyle = 'Hidden'

# Stop button
$stopButton = New-Object System.Windows.Forms.ToolStripButton
$stopButton.Text = "Stop"
$stopButton.Enabled = $false
$stopButton.add_Click({ Stop-CurrentProcess })
$toolStrip.Items.Add($stopButton)

# Separator
$toolStrip.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# Clear button
$clearButton = New-Object System.Windows.Forms.ToolStripButton
$clearButton.Text = "Clear Output"
$clearButton.add_Click({ $richTextBox.Clear(); Update-Status "Output cleared" })
$toolStrip.Items.Add($clearButton)

# Add separator and spacer to push status to right
$toolStrip.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# Create a flexible space using a ToolStripLabel
$spacer = New-Object System.Windows.Forms.ToolStripLabel
$spacer.Text = ""
$spacer.AutoSize = $false
$spacer.Width = 200
$toolStrip.Items.Add($spacer)

# Status label
$statusLabel = New-Object System.Windows.Forms.ToolStripLabel
$statusLabel.Text = "Ready"
$statusLabel.Alignment = 'Right'
$toolStrip.Items.Add($statusLabel)

# Create main container
$mainContainer = New-Object System.Windows.Forms.TableLayoutPanel
$mainContainer.Dock = 'Fill'
$mainContainer.ColumnCount = 1
$mainContainer.RowCount = 2
$mainContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
$mainContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))

# Panel for buttons with improved layout
$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.AutoSize = $true
$buttonPanel.WrapContents = $true
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$buttonPanel.FlowDirection = 'LeftToRight'
$buttonPanel.BackColor = [System.Drawing.Color]::LightGray

# Generate buttons dynamically with improved styling
foreach ($script in $GitHubScripts) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $script.Description
    $btn.Size = New-Object System.Drawing.Size(180, 35)
    $btn.Margin = New-Object System.Windows.Forms.Padding(5)
    $btn.UseVisualStyleBackColor = $true
    $btn.FlatStyle = 'System'
    $btn.Tag = $script
    $btn.add_Click({
        $scriptInfo = $this.Tag
        Start-Script $scriptInfo.ScriptUrl $scriptInfo.Description
    })
    $buttonPanel.Controls.Add($btn)
}

# Exit button with improved styling
$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text = 'Exit Application'
$exitBtn.Size = New-Object System.Drawing.Size(180, 35)
$exitBtn.Margin = New-Object System.Windows.Forms.Padding(5)
$exitBtn.BackColor = [System.Drawing.Color]::LightCoral
$exitBtn.FlatStyle = 'System'
$exitBtn.add_Click({ $form.Close() })
$buttonPanel.Controls.Add($exitBtn)

# RichTextBox for output with improved formatting
$richTextBox = New-Object System.Windows.Forms.RichTextBox
$richTextBox.Dock = 'Fill'
$richTextBox.ReadOnly = $true
$richTextBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$richTextBox.BackColor = [System.Drawing.Color]::Black
$richTextBox.ForeColor = [System.Drawing.Color]::White
$richTextBox.WordWrap = $true
$richTextBox.ScrollBars = 'Vertical'

# Add controls to container
$mainContainer.Controls.Add($buttonPanel, 0, 0)
$mainContainer.Controls.Add($richTextBox, 0, 1)

# Add all controls to form
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($mainContainer)
$form.Controls.Add($toolStrip)
$form.Controls.Add($menuStrip)

# Initial status
Write-Output "AutoBytePro GUI v2.0 - Ready" "Green"
Write-Output "Select a script from the buttons above to begin execution." "Gray"
Write-Output "Output will appear here in real-time." "Gray"
Update-Status "Ready"

# Show the form
try {
    # Run the application
    [System.Windows.Forms.Application]::Run($form)
} catch {
    Write-Host "Error running application: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
} finally {
    # Cleanup
    Stop-CurrentProcess
    Remove-TempFiles
}