# AutoBytePro GUI – Fixed Complete Version
# Requires PowerShell 5.1 or later
# Author: Earl “deep1ne” Daniels (enhanced)

# Load WinForms & Drawing assemblies, then initialize Application
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try {
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
} catch {
    # Ignore if WinForms was already loaded by the host
}

# Global state
$script:currentProcess    = $null
$script:tempFiles         = @()
$script:WriteLogDelegate  = $null

# Cross-scope Write-Log function
function Write-Log { param($m) & $script:WriteLogDelegate $m }

# Creates a cross-thread-safe delegate for writing to a RichTextBox (no SelectionStart calls)
function New-WriteLogDelegate {
    param([System.Windows.Forms.RichTextBox]$TextBox)
    $handler = {
        param($msg)
        if ($TextBox.InvokeRequired) {
            $TextBox.Invoke($script:WriteLogDelegate, $msg)
            return
        }
        $timestamp = (Get-Date).ToString('HH:mm:ss')
        $TextBox.AppendText("$timestamp  $msg`r`n")
        $TextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    return [Action[string]]$handler
}

# Builds and returns the main Form
function Start-AutoByteProGUI {
    # Main form
    $form = [System.Windows.Forms.Form]::new()
    $form.Text          = 'AutoBytePro GUI v2.0'
    $form.Size          = [System.Drawing.Size]::new(1000,700)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize   = [System.Drawing.Size]::new(800,600)
    $form.Icon          = [System.Drawing.SystemIcons]::Application

    # Output box
    $richTextBox = [System.Windows.Forms.RichTextBox]::new()
    $richTextBox.Dock       = 'Fill'
    $richTextBox.ReadOnly   = $true
    $richTextBox.Font       = [System.Drawing.Font]::new('Consolas',10)
    $richTextBox.BackColor  = [System.Drawing.Color]::Black
    $richTextBox.ForeColor  = [System.Drawing.Color]::White
    $richTextBox.WordWrap   = $true
    $richTextBox.ScrollBars = 'Vertical'

    # Install Write-Log delegate
    $script:WriteLogDelegate = New-WriteLogDelegate -TextBox $richTextBox

    # Status bar
    $script:statusLabel = [System.Windows.Forms.ToolStripStatusLabel]::new('Ready')
    $statusStrip = [System.Windows.Forms.StatusStrip]::new()
    $statusStrip.Items.Add($script:statusLabel) | Out-Null

    # Toolbar with Stop & Clear
    $toolStrip  = [System.Windows.Forms.ToolStrip]::new()
    $script:stopButton = [System.Windows.Forms.ToolStripButton]::new('Stop')
    $script:stopButton.Enabled = $false
    $script:stopButton.add_Click({ Stop-CurrentProcess })
    $clearButton = [System.Windows.Forms.ToolStripButton]::new('Clear Output')
    $clearButton.add_Click({
        $richTextBox.Clear()
        $script:statusLabel.Text = 'Output cleared'
    })
    $toolStrip.Items.AddRange(@(
        $script:stopButton,
        [System.Windows.Forms.ToolStripSeparator]::new(),
        $clearButton
    ))

    # Panel for script buttons
    $script:buttonPanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $script:buttonPanel.Dock         = 'Top'
    $script:buttonPanel.AutoSize     = $true
    $script:buttonPanel.Padding      = [System.Windows.Forms.Padding]::new(10)
    $script:buttonPanel.BackColor    = [System.Drawing.Color]::LightGray

    # Define GitHub-hosted scripts
    $GitHubScripts = @(
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1";           Text = "Disk Cleaner" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1";   Text = "Enable Files On Demand" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadandInstallPackage.ps1"; Text = "Download & Install Package" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckUserProfileIssue.ps1"; Text = "Check User Profile" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1";       Text = "Dell Bloatware Remover" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1";    Text = "Reset & Install Windows Update" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1";     Text = "Windows System Repair" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Text = "Reset Windows Search DB" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallMSProjects.ps1";       Text = "Install MS Projects" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1";         Text = "Check Drive Space" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1";       Text = "Internet Speed Test" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1";     Text = "Internet Latency Test" },
        @{ Url = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WorkPaperMonitorTroubleShooter.ps1"; Text = "Monitor Troubleshooter" }
    )
    foreach ($s in $GitHubScripts) {
        $btn = [System.Windows.Forms.Button]::new()
        $btn.Text   = $s.Text
        $btn.Size   = [System.Drawing.Size]::new(180,35)
        $btn.Margin = [System.Windows.Forms.Padding]::new(5)
        $btn.Tag    = $s
        $btn.add_Click({ Start-Script -Url $this.Tag.Url -Description $this.Tag.Text })
        $script:buttonPanel.Controls.Add($btn)
    }
    $exitBtn = [System.Windows.Forms.Button]::new('Exit')
    $exitBtn.Size      = [System.Drawing.Size]::new(180,35)
    $exitBtn.Margin    = [System.Windows.Forms.Padding]::new(5)
    $exitBtn.BackColor = [System.Drawing.Color]::LightCoral
    $exitBtn.add_Click({ $form.Close() })
    $script:buttonPanel.Controls.Add($exitBtn)

    # Layout container
    $layout = [System.Windows.Forms.TableLayoutPanel]::new()
    $layout.Dock      = 'Fill'
    $layout.RowCount  = 2
    $layout.Controls.Add($script:buttonPanel,    0,0)
    $layout.Controls.Add($richTextBox,     0,1)

    # Assemble form
    $form.Controls.Add($layout)
    $form.Controls.Add($toolStrip)
    $form.Controls.Add($statusStrip)

    Write-Log 'AutoBytePro GUI v2.0 – Ready'
    Write-Log 'Select a script to begin execution.'

    return $form
}
function Start-Script {
    param(
        [Parameter(Mandatory)] [string] $Url,
        [Parameter(Mandatory)] [string] $Description
    )
    $script:buttonPanel.Enabled = $false
    $script:stopButton.Enabled  = $true
    $script:statusLabel.Text    = "Running: $Description"

    Write-Log "▶ $Description"
    Write-Log "Downloading from $Url"

    try {
        $wc      = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent','AutoBytePro/2.0')
        $content = $wc.DownloadString($Url)
        $wc.Dispose()
        if (-not $content) { throw 'Downloaded script is empty.' }

        $file = Join-Path $env:TEMP ("ABP_{0}.ps1" -f [guid]::NewGuid())
        $content | Out-File -FilePath $file -Encoding UTF8
        $script:tempFiles += $file
        Write-Log "Saved to $file"

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

        $script:currentProcess.add_OutputDataReceived({
            param($s,$e)
            if ($e.Data) { Write-Log $e.Data }
        })
        $script:currentProcess.add_ErrorDataReceived({
            param($s,$e)
            if ($e.Data) { Write-Log "ERROR: $($e.Data)" }
        })
        $script:currentProcess.add_Exited({
            param($s,$e)
            $code = $script:currentProcess.ExitCode
            $script:buttonPanel.Enabled = $true
            $script:stopButton.Enabled  = $false
            $script:statusLabel.Text    = if ($code -eq 0) { 'Ready' } else { "Error ($code)" }
        })
        $script:currentProcess.Start() | Out-Null
        $script:currentProcess.BeginOutputReadLine()
        $script:currentProcess.BeginErrorReadLine()
    } catch {
        Write-Log "ERROR: $_"
        $script:buttonPanel.Enabled = $true
        $script:stopButton.Enabled  = $false
        $script:statusLabel.Text    = 'Error'
    }
}

function Stop-CurrentProcess {
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        try {
            $script:currentProcess.Kill()
            $script:currentProcess.WaitForExit()
            Write-Log "Process terminated by user."
        } catch {
            Write-Log "Failed to terminate process: $_"
        } finally {
            $script:buttonPanel.Enabled = $true
            $script:stopButton.Enabled  = $false
            $script:statusLabel.Text    = 'Stopped'
        }
    }
}
