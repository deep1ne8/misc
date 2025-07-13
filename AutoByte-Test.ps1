# AutoBytePro GUI – Fixed Complete Version
# Requires PowerShell 5.1 or later
# Author: Earl “deep1ne” Daniels (enhanced)

# Load WinForms & Drawing assemblies, then initialize Application
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# Global state
$script:currentProcess      = $null
$script:tempFiles          = @()
$script:WriteLogDelegate   = $null

# Creates a cross-thread‐safe delegate for writing to a RichTextBox
function New-WriteLogDelegate {
    param(
        [System.Windows.Forms.RichTextBox] $TextBox
    )
    $handler = {
        param($msg, $clr)
        if ($TextBox.InvokeRequired) {
            $TextBox.Invoke($script:WriteLogDelegate, $msg, $clr)
            return
        }
        $TextBox.SelectionStart  = $TextBox.TextLength
        $TextBox.SelectionLength = 0
        $TextBox.SelectionColor  = [System.Drawing.Color]::FromName($clr)
        $TextBox.AppendText((Get-Date -Format "HH:mm:ss") + "  " + $msg + "`r`n")
        $TextBox.SelectionColor  = [System.Drawing.Color]::Black
        $TextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    return [Action[string,string]]$handler
}

# Main GUI builder
function Start-AutoByteProGUI {
    # --- Form and Controls ---
    $form = [System.Windows.Forms.Form]::new()
    $form.Text            = 'AutoBytePro GUI v2.0'
    $form.Size            = [System.Drawing.Size]::new(1000,700)
    $form.StartPosition   = 'CenterScreen'
    $form.MinimumSize     = [System.Drawing.Size]::new(800,600)
    $form.Icon            = [System.Drawing.SystemIcons]::Application

    # Output box
    $richTextBox = [System.Windows.Forms.RichTextBox]::new()
    $richTextBox.Dock       = 'Fill'
    $richTextBox.ReadOnly   = $true
    $richTextBox.Font       = [System.Drawing.Font]::new('Consolas',10)
    $richTextBox.BackColor  = [System.Drawing.Color]::Black
    $richTextBox.ForeColor  = [System.Drawing.Color]::White
    $richTextBox.WordWrap   = $true
    $richTextBox.ScrollBars = 'Vertical'

    # Install our single Write-Log function
    $script:WriteLogDelegate = New-WriteLogDelegate -TextBox $richTextBox
    function Write-Log {
        param([string]$Message, [string]$Color = 'Black')
        & $script:WriteLogDelegate $Message $Color
    }

    # Status bar
    $statusLabel = [System.Windows.Forms.ToolStripStatusLabel]::new("Ready")
    $statusStrip = [System.Windows.Forms.StatusStrip]::new()
    $statusStrip.Items.Add($statusLabel) | Out-Null

    # Top toolbar
    $toolStrip  = [System.Windows.Forms.ToolStrip]::new()
    $stopButton = [System.Windows.Forms.ToolStripButton]::new("Stop")
    $stopButton.Enabled = $false
    $stopButton.add_Click({ Stop-CurrentProcess })
    $clearButton = [System.Windows.Forms.ToolStripButton]::new("Clear Output")
    $clearButton.add_Click({ 
        $richTextBox.Clear()
        $statusLabel.Text = "Output cleared"
    })
    $toolStrip.Items.AddRange(@(
        $stopButton,
        [System.Windows.Forms.ToolStripSeparator]::new(),
        $clearButton
    ))

    # Script-button panel
    $buttonPanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $buttonPanel.Dock        = 'Top'
    $buttonPanel.AutoSize    = $true
    $buttonPanel.WrapContents= $true
    $buttonPanel.Padding     = [System.Windows.Forms.Padding]::new(10)
    $buttonPanel.BackColor   = [System.Drawing.Color]::LightGray

    # Define your GitHub scripts
    $GitHubScripts = @(
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1";      Text="Disk Cleaner" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1";Text="Enable Files On Demand" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadandInstallPackage.ps1"; Text="Download & Install Package" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckUserProfileIssue.ps1"; Text="Check User Profile" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1";    Text="Dell Bloatware Remover" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1"; Text="Reset & Install Windows Update" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1"; Text="Windows System Repair" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Text="Reset Windows Search DB" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallMSProjects.ps1"; Text="Install MS Projects" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1"; Text="Check Drive Space" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1"; Text="Internet Speed Test" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1"; Text="Internet Latency Test" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WorkPaperMonitorTroubleShooter.ps1"; Text="Monitor Troubleshooter" }
    )

    # Build buttons
    foreach ($s in $GitHubScripts) {
        $btn = [System.Windows.Forms.Button]::new()
        $btn.Text   = $s.Text
        $btn.Size   = [System.Drawing.Size]::new(180,35)
        $btn.Margin = [System.Windows.Forms.Padding]::new(5)
        $btn.Tag    = $s
        $btn.add_Click({
            $info = $this.Tag
            Start-Script -Url $info.Url -Description $info.Text
        })
        $buttonPanel.Controls.Add($btn)
    }

    # Exit button
    $exitBtn = [System.Windows.Forms.Button]::new("Exit Application")
    $exitBtn.Size      = [System.Drawing.Size]::new(180,35)
    $exitBtn.Margin    = [System.Windows.Forms.Padding]::new(5)
    $exitBtn.BackColor = [System.Drawing.Color]::LightCoral
    $exitBtn.add_Click({ $form.Close() })
    $buttonPanel.Controls.Add($exitBtn)

    # Layout
    $main = [System.Windows.Forms.TableLayoutPanel]::new()
    $main.Dock     = 'Fill'
    $main.RowCount = 2
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent',100)))
    $main.Controls.Add($buttonPanel, 0, 0)
    $main.Controls.Add($richTextBox,  0, 1)

    # Assemble form
    $form.Controls.Add($main)
    $form.Controls.Add($toolStrip)
    $form.Controls.Add($statusStrip)

    Write-Log "AutoBytePro GUI v2.0 – Ready" "Green"
    Write-Log "Select a script to begin execution."      "Gray"

    return $form
}

# Starts, monitors, and logs a script run
function Start-Script {
    param(
        [Parameter(Mandatory)] [string] $Url,
        [Parameter(Mandatory)] [string] $Description
    )

    # disable UI
    $buttonPanel.Enabled = $false
    $stopButton.Enabled  = $true
    $statusLabel.Text    = "Running: $Description"

    Write-Log "┌─────────────────────────────────────────" "Blue"
    Write-Log "▶ $Description"                             "Blue"
    Write-Log "Downloading from $Url"                      "Green"

    try {
        # Download
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent","AutoBytePro/2.0")
        $content = $wc.DownloadString($Url)
        $wc.Dispose()
        if (-not $content) { throw "Empty script content." }

        # Save to temp
        $file = Join-Path $env:TEMP ("ABP_{0}.ps1" -f ([guid]::NewGuid()))
        $content | Out-File -FilePath $file -Encoding UTF8
        $script:tempFiles += $file
        Write-Log "Saved to $file" "Gray"

        # Configure process
        $psi = [System.Diagnostics.ProcessStartInfo]::new(
            'powershell.exe',
            "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$file`""
        )
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true

        $script:currentProcess = [System.Diagnostics.Process]::new()
        $script:currentProcess.StartInfo           = $psi
        $script:currentProcess.EnableRaisingEvents = $true

        # Wire events
        $script:currentProcess.add_OutputDataReceived({
            param($s,$e)
            if ($e.Data) { Write-Log $e.Data "Black" }
        })
        $script:currentProcess.add_ErrorDataReceived({
            param($s,$e)
            if ($e.Data) { Write-Log "ERROR: $($e.Data)" "Red" }
        })
        $script:currentProcess.add_Exited({
            param($s,$e)
            $code = $script:currentProcess.ExitCode
            Write-Log "└─────────────────────────────────────────" "Blue"
            if ($code -eq 0) {
                Write-Log "✔ Completed ($code)" "Green"
                $statusLabel.Text = "Ready"
            } else {
                Write-Log "✖ Exit Code: $code" "Orange"
                $statusLabel.Text = "Error ($code)"
            }
            $buttonPanel.Enabled = $true
            $stopButton.Enabled  = $false
        })

        # Launch
        $script:currentProcess.Start()            | Out-Null
        $script:currentProcess.BeginOutputReadLine()
        $script:currentProcess.BeginErrorReadLine()
    }
    catch {
        Write-Log "CRITICAL: $_" "Red"
        $buttonPanel.Enabled = $true
        $stopButton.Enabled  = $false
        $statusLabel.Text    = "Error"
    }
}

# Allows the user to stop the running process
function Stop-CurrentProcess {
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        try { $script:currentProcess.Kill() }
        catch { Write-Log "Failed to terminate: $_" "Red" }
        finally {
            $buttonPanel.Enabled = $true
            $stopButton.Enabled  = $false
            $statusLabel.Text    = "Stopped"
        }
    }
}

# Run the GUI
[System.Windows.Forms.Application]::Run( (Start-AutoByteProGUI) )

# Cleanup any leftover temp files
foreach ($f in $script:tempFiles) {
    if (Test-Path $f) {
        Remove-Item $f -Force -ErrorAction SilentlyContinue
    }
}
