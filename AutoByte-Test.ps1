# AutoBytePro GUI – v3.9 Full Script with Correct Control Scoping
# Requires PowerShell 5.1 or later
# Author: Earl "deep1ne" Daniels

# 1) Load WinForms BEFORE any controls
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

# 2) State placeholders (script scope)
$richTextBox    = $null
$statusLabel    = $null
$buttonPanel    = $null
$stopButton     = $null
$currentProcess = $null
$tempFiles      = @()

# 3) Approved-verb logging
function Write-Log {
    param([string]$Message)
    if ($richTextBox) {
        if ($richTextBox.InvokeRequired) {
            $richTextBox.Invoke([Action[string]]{ param($m) Write-Log $m }, $Message)
        } else {
            $richTextBox.AppendText("$((Get-Date).ToString('HH:mm:ss'))  $Message`r`n")
            $richTextBox.ScrollToCaret()
        }
    } else {
        Write-Host $Message
    }
}

# 4) Approved-verb process termination
function Stop-ProcessExecution {
    if ($currentProcess -and -not $currentProcess.HasExited) {
        try {
            $currentProcess.Kill()
            $currentProcess.WaitForExit()
        } catch {
            Write-Log "Failed to terminate: $_"
        } finally {
            if ($currentProcess.HasExited) { $currentProcess.Dispose() }
            $buttonPanel.Enabled = $true
            $stopButton.Enabled  = $false
            $statusLabel.Text    = 'Stopped'
        }
    }
}

# 5) Approved-verb script invocation with timer-based UI pumping
function Invoke-RemoteScript {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Description
    )
    # Disable UI
    $buttonPanel.Enabled = $false
    $stopButton.Enabled  = $true
    $statusLabel.Text    = "Running: $Description"
    Write-Log "▶ $Description"
    Write-Log "Downloading from $Url"

    try {
        # Download script
        $wc         = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent','AutoBytePro/3.9')
        $scriptText = $wc.DownloadString($Url)
        $wc.Dispose()
        if (-not $scriptText) { throw 'Downloaded script is empty.' }

        # Save temporary file
        $tempFile = Join-Path $env:TEMP ("ABP_{0}.ps1" -f [guid]::NewGuid())
        $scriptText | Out-File -FilePath $tempFile -Encoding UTF8
        $tempFiles += $tempFile
        Write-Log "Saved to $tempFile"

        # Configure process
        $psi = New-Object System.Diagnostics.ProcessStartInfo('powershell.exe',
            "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$tempFile`"")
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true

        # Start process
        $currentProcess = New-Object System.Diagnostics.Process
        $currentProcess.StartInfo           = $psi
        $currentProcess.EnableRaisingEvents = $true
        $currentProcess.Start() | Out-Null

        # Promote streams to script scope
        Set-Variable -Name stdout -Scope Script -Value $currentProcess.StandardOutput
        Set-Variable -Name stderr -Scope Script -Value $currentProcess.StandardError

        # Capture UI controls locally
        $rtb      = $richTextBox
        $btnPanel = $buttonPanel
        $stBtn    = $stopButton
        $stLbl    = $statusLabel

        # Timer to pump streams on UI thread
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 200
        $timer.Add_Tick({
            try {
                # STDOUT
                if ($null -ne $script:stdout) {
                    while ($script:stdout.Peek() -ge 0) {
                        $line = $script:stdout.ReadLine()
                        if ($rtb.InvokeRequired) {
                            $rtb.Invoke([Action]{ $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  $line`r`n"); $rtb.ScrollToCaret() })
                        } else {
                            $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  $line`r`n")
                            $rtb.ScrollToCaret()
                        }
                    }
                }
                # STDERR
                if ($null -ne $script:stderr) {
                    while ($script:stderr.Peek() -ge 0) {
                        $line = $script:stderr.ReadLine()
                        if ($rtb.InvokeRequired) {
                            $rtb.Invoke([Action]{ $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  ERROR: $line`r`n"); $rtb.ScrollToCaret() })
                        } else {
                            $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  ERROR: $line`r`n")
                            $rtb.ScrollToCaret()
                        }
                    }
                }
                # On exit
                if ($currentProcess.HasExited) {
                    $timer.Stop()
                    $exitCode = $currentProcess.ExitCode
                    # UI update
                    $updateUI = {
                        if ($exitCode -eq 0) { $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  ✔ Completed ($exitCode)`r`n") }
                        else             { $rtb.AppendText("$((Get-Date).ToString('HH:mm:ss'))  ✖ Exit ($exitCode)`r`n") }
                        $rtb.ScrollToCaret()
                        $btnPanel.Enabled = $true
                        $stBtn.Enabled   = $false
                        $stLbl.Text      = if ($exitCode -eq 0) { 'Ready' } else { "Error ($exitCode)" }
                    }
                    if ($rtb.InvokeRequired) { $rtb.Invoke([Action]$updateUI) } else { & $updateUI }

                    # Dispose process
                    $currentProcess.Dispose()
                }
            } catch {
                Write-Log "Timer error: $_"
            }
        })
        $timer.Start()
    }
    catch {
        Write-Log "CRITICAL: $_"
        $buttonPanel.Enabled = $true
        $stopButton.Enabled  = $false
        $statusLabel.Text    = 'Error'
    }
}

# 6) Build and return the Form
function Start-AutoByteProGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'AutoBytePro GUI v3.9'
    $form.Size          = New-Object System.Drawing.Size(1000,700)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize   = New-Object System.Drawing.Size(800,600)
    $form.Icon          = [System.Drawing.SystemIcons]::Application

    # Clean up temp files on close
    $form.Add_FormClosing({ foreach ($f in $tempFiles) { if (Test-Path $f) { Remove-Item $f -Force } } })

    # Output box
    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Dock       = 'Fill'
    $richTextBox.ReadOnly   = $true
    $richTextBox.Font       = New-Object System.Drawing.Font('Consolas',10)
    $richTextBox.BackColor  = [System.Drawing.Color]::Black
    $richTextBox.ForeColor  = [System.Drawing.Color]::White
    $richTextBox.WordWrap   = $true
    $richTextBox.ScrollBars = 'Vertical'

    # Status strip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel 'Ready'
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    [void]$statusStrip.Items.Add($statusLabel)

    # Toolbar
    $toolStrip = New-Object System.Windows.Forms.ToolStrip
    $stopButton = New-Object System.Windows.Forms.ToolStripButton 'Stop'
    $stopButton.Enabled = $false
    $stopButton.Add_Click({ Stop-ProcessExecution })
    $clearButton = New-Object System.Windows.Forms.ToolStripButton 'Clear'
    $clearButton.Add_Click({ $richTextBox.Clear(); $statusLabel.Text = 'Output cleared' })
    [void]$toolStrip.Items.AddRange(@($stopButton, [System.Windows.Forms.ToolStripSeparator]::new(), $clearButton))

    # Button panel
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock      = 'Top'
    $buttonPanel.AutoSize  = $true
    $buttonPanel.Padding   = New-Object System.Windows.Forms.Padding(10)
    $buttonPanel.BackColor = [System.Drawing.Color]::LightGray

        # 13 GitHub scripts
    $GitHubScripts = @(
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1";           Text="Disk Cleaner" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1";   Text="Enable Files On Demand" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadandInstallPackage.ps1"; Text="Download & Install Package" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckUserProfileIssue.ps1";   Text="Check User Profile" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1";       Text="Dell Bloatware Remover" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1";    Text="Reset & Install Windows Update" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1";     Text="Windows System Repair" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Text="Reset Windows Search DB" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallMSProjects.ps1";       Text="Install MS Projects" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1";         Text="Check Drive Space" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1";       Text="Internet Speed Test" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1";     Text="Internet Latency Test" },
        @{ Url="https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WorkPaperMonitorTroubleShooter.ps1"; Text="Monitor Troubleshooter" }
    )

    # Create buttons for each script
    foreach ($s in $GitHubScripts) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text   = $s.Text
        $btn.Size   = New-Object System.Drawing.Size(180,35)
        $btn.Margin = New-Object System.Windows.Forms.Padding(5)
        $sb = [scriptblock]::Create("Invoke-RemoteScript -Url '$($s.Url)' -Description '$($s.Text)'")
        $btn.Add_Click($sb)
        [void]$buttonPanel.Controls.Add($btn)
    

    # Exit button
    $exitBtn = New-Object System.Windows.Forms.Button
    $exitBtn.Text   = 'Exit Application'
    $exitBtn.Size   = New-Object System.Drawing.Size(180,35)
    $exitBtn.Margin = New-Object System.Windows.Forms.Padding(5)
    $exitBtn.Add_Click({ $form.Close() })
    [void]$buttonPanel.Controls.Add($exitBtn)

    # Layout
    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock        = 'Fill'
    $layout.ColumnCount = 1
    [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent',100)))
    $layout.RowCount    = 2
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent',100)))
    [void]$layout.Controls.Add($buttonPanel, 0, 0)
    [void]$layout.Controls.Add($richTextBox, 0, 1)

    # Assemble
    [void]$form.Controls.Add($layout)
    [void]$form.Controls.Add($toolStrip)
    [void]$form.Controls.Add($statusStrip)

    Write-Log 'AutoBytePro GUI v3.9 Ready'
    Write-Log 'Select a button to run a script.'
    return $form
    }
}

# 7) Launch
$form = Start-AutoByteProGUI
[System.Windows.Forms.Application]::Run($form)
