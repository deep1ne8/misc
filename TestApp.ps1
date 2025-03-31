# Optimized Fix for Show-RemoteExecutionSettings Function
function Show-RemoteExecutionSettings {
    [CmdletBinding()]
    [OutputType([void])]
    param()

    # Ensure $script:Colors are properly initialized
    if (-not $script:Colors) {
        $script:Colors = @{
            Info      = "Cyan"
            Primary   = "Yellow"
            Success   = "Green"
            Error     = "Red"
            Warning   = "Magenta"
            Accent    = "Blue"
            Secondary = "Gray"
        }
    }

    Show-Header -Title "Remote Execution Settings"

    Write-Host "Configure remote execution targets and options." -ForegroundColor $script:Colors.Info
    Write-Host ""

    # Current settings display
    Write-Host "CURRENT SETTINGS" -ForegroundColor $script:Colors.Primary
    $enabledText = if ($script:Config.RemoteExecution.Enabled) { "Enabled" } else { "Disabled" }
    $enabledColor = if ($script:Config.RemoteExecution.Enabled) { $script:Colors.Success } else { $script:Colors.Error }

    Write-Host "  Status:                $enabledText" -ForegroundColor $enabledColor
    Write-Host "  Authentication Method: $($script:Config.RemoteExecution.AuthenticationMethod)" -ForegroundColor $script:Colors.Info
    Write-Host "  Max Concurrent:        $($script:Config.RemoteExecution.MaxConcurrentSessions)" -ForegroundColor $script:Colors.Info
    Write-Host "  Timeout:               $($script:Config.RemoteExecution.Timeout) seconds" -ForegroundColor $script:Colors.Info

    # Display current targets
    Write-Host ""
    Write-Host "CURRENT TARGETS" -ForegroundColor $script:Colors.Primary

    if ($script:RemoteTargets.Count -gt 0) {
        foreach ($target in $script:RemoteTargets) {
            $isReachable = Test-Connection -ComputerName $target -Count 1 -Quiet
            $statusColor = if ($isReachable) { $script:Colors.Success } else { $script:Colors.Error }
            $statusText = if ($isReachable) { "Online" } else { "Offline" }

            Write-Host "  $target" -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host " - $statusText" -ForegroundColor $statusColor
        }
    }
    else {
        Write-Host "  No remote targets configured." -ForegroundColor $script:Colors.Warning
    }

    # Menu options
    Write-Host ""
    Write-Host "OPTIONS" -ForegroundColor $script:Colors.Primary
    Write-Host "1. Toggle Remote Execution (Currently: $enabledText)" -ForegroundColor $script:Colors.Success
    Write-Host "2. Add Remote Target" -ForegroundColor $script:Colors.Success
    Write-Host "3. Remove Remote Target" -ForegroundColor $script:Colors.Success
    Write-Host "4. Test Remote Connectivity" -ForegroundColor $script:Colors.Success
    Write-Host "5. Change Authentication Method" -ForegroundColor $script:Colors.Success
    Write-Host "6. Update Max Concurrent Sessions" -ForegroundColor $script:Colors.Success
    Write-Host "B. Back to Main Menu" -ForegroundColor $script:Colors.Warning

    Show-Footer

    # Ensure Show-Prompt is correctly implemented
    function Show-Prompt {
        param (
            [string]$Message,
            [string[]]$Options
        )
        do {
            Write-Host $Message -ForegroundColor Cyan
            $inputFromUser = Read-Host
            if ($Options -contains $inputFromUser) {
                return $inputFromUser
            } else {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            }
        } while ($true)
    }

    $choice = Show-Prompt -Message "Enter your choice" -Options @("1", "2", "3", "4", "5", "6", "B")

    switch ($choice) {
        "1" {
            # Toggle remote execution
            $script:Config.RemoteExecution.Enabled = -not $script:Config.RemoteExecution.Enabled
            $newStatus = if ($script:Config.RemoteExecution.Enabled) { "enabled" } else { "disabled" }
            Write-Host "Remote execution is now $newStatus." -ForegroundColor $script:Colors.Success
            Export-Configuration -Configuration $script:Config -Path $ConfigPath
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "2" {
            # Add remote target
            $newTarget = Show-Prompt -Message "Enter computer name or IP address"

            if ([string]::IsNullOrWhiteSpace($newTarget)) {
                Write-Host "Invalid target name." -ForegroundColor $script:Colors.Error
            }
            elseif ($script:RemoteTargets -contains $newTarget) {
                Write-Host "Target already exists in the list." -ForegroundColor $script:Colors.Warning
            }
            else {
                # Test connectivity
                Write-Host "Testing connectivity to $newTarget..." -ForegroundColor $script:Colors.Info
                $isReachable = Test-Connection -ComputerName $newTarget -Count 1 -Quiet

                if ($isReachable) {
                    $script:RemoteTargets += $newTarget
                    $script:Config.RemoteExecution.DefaultTargets += $newTarget
                    Write-Host "Target $newTarget added successfully." -ForegroundColor $script:Colors.Success
                    Export-Configuration -Configuration $script:Config -Path $ConfigPath
                }
                else {
                    $confirm = Show-Prompt -Message "Target $newTarget is not reachable. Add anyway? (Y/N)" -Options @("Y", "N")

                    if ($confirm -eq "Y") {
                        $script:RemoteTargets += $newTarget
                        $script:Config.RemoteExecution.DefaultTargets += $newTarget
                        Write-Host "Target $newTarget added (offline)." -ForegroundColor $script:Colors.Warning
                        Export-Configuration -Configuration $script:Config -Path $ConfigPath
                    }
                }
            }

            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "3" {
            # Remove remote target
            if ($script:RemoteTargets.Count -eq 0) {
                Write-Host "No targets to remove." -ForegroundColor $script:Colors.Warning
                Start-Sleep -Seconds 1
                Show-RemoteExecutionSettings
                return
            }

            Write-Host "Select a target to remove:" -ForegroundColor $script:Colors.Info

            for ($i = 0; $i -lt $script:RemoteTargets.Count; $i++) {
                Write-Host "$($i+1). $($script:RemoteTargets[$i])" -ForegroundColor $script:Colors.Success
            }

            $targetChoice = Show-Prompt -Message "Enter target number to remove (or C to cancel)" -Options @("C") + (1..$script:RemoteTargets.Count | ForEach-Object { "$_" })

            if ($targetChoice -eq "C") {
                Show-RemoteExecutionSettings
                return
            }

            $parsedChoice = $null
            if ([int]::TryParse($targetChoice, [ref]$parsedChoice) -and $parsedChoice -gt 0 -and $parsedChoice -le $script:RemoteTargets.Count) {
                $targetToRemove = $script:RemoteTargets[$parsedChoice - 1]
                $script:RemoteTargets = $script:RemoteTargets | Where-Object { $_ -ne $targetToRemove }
                $script:Config.RemoteExecution.DefaultTargets = $script:Config.RemoteExecution.DefaultTargets | Where-Object { $_ -ne $targetToRemove }

                Write-Host "Target $targetToRemove removed." -ForegroundColor $script:Colors.Success
                Export-Configuration -Configuration $script:Config -Path $ConfigPath
            }
            else {
                Write-Host "Invalid selection." -ForegroundColor $script:Colors.Error
            }

            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "4" {
            # Test connectivity
            if ($script:RemoteTargets.Count -eq 0) {
                Write-Host "No targets to test." -ForegroundColor $script:Colors.Warning
                Start-Sleep -Seconds 1
                Show-RemoteExecutionSettings
                return
            }

            Write-Host "Testing connectivity to all targets..." -ForegroundColor $script:Colors.Info

            foreach ($target in $script:RemoteTargets) {
                $isReachable = Test-Connection -ComputerName $target -Count 1 -Quiet
                $statusColor = if ($isReachable) { $script:Colors.Success } else { $script:Colors.Error }
                $statusText = if ($isReachable) { "Reachable" } else { "Unreachable" }

                Write-Host "  $target - $statusText" -ForegroundColor $statusColor

                if ($isReachable) {
                    # Test PowerShell remoting
                    try {
                        Invoke-Command -ComputerName $target -ScriptBlock { "OK" } -ErrorAction Stop | Out-Null
                        Write-Host "    PowerShell Remoting: " -NoNewline -ForegroundColor $script:Colors.Info
                        Write-Host "Enabled" -ForegroundColor $script:Colors.Success
                    }
                    catch {
                        Write-Host "    PowerShell Remoting: " -NoNewline -ForegroundColor $script:Colors.Info
                        Write-Host "Disabled or Error" -ForegroundColor $script:Colors.Error
                        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor $script:Colors.Error
                    }
                }
            }

            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor $script:Colors.Warning
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            Show-RemoteExecutionSettings
        }
        "5" {
            # Change authentication method
            Write-Host "Select authentication method:" -ForegroundColor $script:Colors.Info
            Write-Host "1. Credential (username/password)" -ForegroundColor $script:Colors.Success
            Write-Host "2. Certificate" -ForegroundColor $script:Colors.Success
            Write-Host "3. CredSSP (enables credential delegation)" -ForegroundColor $script:Colors.Success

            $authChoice = Show-Prompt -Message "Enter choice" -Options @("1", "2", "3")

            switch ($authChoice) {
                "1" { $script:Config.RemoteExecution.AuthenticationMethod = "Credential" }
                "2" { $script:Config.RemoteExecution.AuthenticationMethod = "Certificate" }
                "3" { $script:Config.RemoteExecution.AuthenticationMethod = "CredSSP" }
            }

            Write-Host "Authentication method updated to $($script:Config.RemoteExecution.AuthenticationMethod)." -ForegroundColor $script:Colors.Success
            Export-Configuration -Configuration $script:Config -Path $ConfigPath

            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "6" {
            # Update max concurrent sessions
            $maxConcurrent = Show-Prompt -Message "Enter maximum concurrent sessions (1-20)" -Options @("1".."20")

            if ($maxConcurrent -match '^\d+$' -and [int]$maxConcurrent -ge 1 -and [int]$maxConcurrent -le 20) {
                $script:Config.RemoteExecution.MaxConcurrentSessions = [int]$maxConcurrent
                Write-Host "Max concurrent sessions updated to $maxConcurrent." -ForegroundColor $script:Colors.Success
                Export-Configuration -Configuration $script:Config -Path $ConfigPath
            }
            else {
                Write-Host "Invalid value. Must be between 1 and 20." -ForegroundColor $script:Colors.Error
            }

            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "B" {
            Show-MainMenu
        }
    }
}
function Show-RemoteExecutionSettings {
    [CmdletBinding()]
    [OutputType([void])]
    param()
    
    Show-Header -Title "Remote Execution Settings"
    
    Write-Host "Configure remote execution targets and options." -ForegroundColor $script:Colors.Info
    Write-Host ""
    
    # Current settings display
    Write-Host "CURRENT SETTINGS" -ForegroundColor $script:Colors.Primary
    $enabledText = if ($script:Config.RemoteExecution.Enabled) { "Enabled" } else { "Disabled" }
    $enabledColor = if ($script:Config.RemoteExecution.Enabled) { $script:Colors.Success } else { $script:Colors.Error }
    
    Write-Host "  Status:                $enabledText" -ForegroundColor $enabledColor
    Write-Host "  Authentication Method: $($script:Config.RemoteExecution.AuthenticationMethod)" -ForegroundColor $script:Colors.Info
    Write-Host "  Max Concurrent:        $($script:Config.RemoteExecution.MaxConcurrentSessions)" -ForegroundColor $script:Colors.Info
    Write-Host "  Timeout:               $($script:Config.RemoteExecution.Timeout) seconds" -ForegroundColor $script:Colors.Info
    
    # Display current targets
    Write-Host ""
    Write-Host "CURRENT TARGETS" -ForegroundColor $script:Colors.Primary
    
    if ($script:RemoteTargets.Count -gt 0) {
        foreach ($target in $script:RemoteTargets) {
            $isReachable = Test-Connection -ComputerName $target -Count 1 -Quiet
            $statusColor = if ($isReachable) { $script:Colors.Success } else { $script:Colors.Error }
            $statusText = if ($isReachable) { "Online" } else { "Offline" }
            
            Write-Host "  $target" -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host " - $statusText" -ForegroundColor $statusColor
        }
    }
    else {
        Write-Host "  No remote targets configured." -ForegroundColor $script:Colors.Warning
    }
    
    # Menu options
    Write-Host ""
    Write-Host "OPTIONS" -ForegroundColor $script:Colors.Primary
    Write-Host "1. Toggle Remote Execution (Currently: $enabledText)" -ForegroundColor $script:Colors.Success
    Write-Host "2. Add Remote Target" -ForegroundColor $script:Colors.Success
    Write-Host "3. Remove Remote Target" -ForegroundColor $script:Colors.Success
    Write-Host "4. Test Remote Connectivity" -ForegroundColor $script:Colors.Success
    Write-Host "5. Change Authentication Method" -ForegroundColor $script:Colors.Success
    Write-Host "6. Update Max Concurrent Sessions" -ForegroundColor $script:Colors.Success
    Write-Host "B. Back to Main Menu" -ForegroundColor $script:Colors.Warning
    
    Show-Footer
    
    function Show-Prompt {
        param (
            [string]$Message,
            [string[]]$Options
        )
        do {
            Write-Host $Message -ForegroundColor $script:Colors.Info
            $inputFromUser = Read-Host
            if ($Options -contains $inputFromUser) {
                return $inputFromUser
            } else {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            }
        } while ($true)
    }

    $choice = Show-Prompt -Message "Enter your choice" -Options @("1", "2", "3", "4", "5", "6", "B")
    
    switch ($choice) {
        "1" {
            # Toggle remote execution
            $script:Config.RemoteExecution.Enabled = -not $script:Config.RemoteExecution.Enabled
            $newStatus = if ($script:Config.RemoteExecution.Enabled) { "enabled" } else { "disabled" }
            Write-Host "Remote execution is now $newStatus." -ForegroundColor $script:Colors.Success
            Export-Configuration -Configuration $script:Config -Path $ConfigPath
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "2" {
            # Add remote target
            $newTarget = Show-Prompt -Message "Enter computer name or IP address"
            
            if ([string]::IsNullOrWhiteSpace($newTarget)) {
                Write-Host "Invalid target name." -ForegroundColor $script:Colors.Error
            }
            elseif ($script:RemoteTargets -contains $newTarget) {
                Write-Host "Target already exists in the list." -ForegroundColor $script:Colors.Warning
            }
            else {
                # Test connectivity
                Write-Host "Testing connectivity to $newTarget..." -ForegroundColor $script:Colors.Info
                $isReachable = Test-Connection -ComputerName $newTarget -Count 1 -Quiet
                
                if ($isReachable) {
                    $script:RemoteTargets += $newTarget
                    $script:Config.RemoteExecution.DefaultTargets += $newTarget
                    Write-Host "Target $newTarget added successfully." -ForegroundColor $script:Colors.Success
                    Export-Configuration -Configuration $script:Config -Path $ConfigPath
                }
                else {
                    $confirm = Show-Prompt -Message "Target $newTarget is not reachable. Add anyway? (Y/N)" -Options @("Y", "N")
                    
                    if ($confirm -eq "Y") {
                        $script:RemoteTargets += $newTarget
                        $script:Config.RemoteExecution.DefaultTargets += $newTarget
                        Write-Host "Target $newTarget added (offline)." -ForegroundColor $script:Colors.Warning
                        Export-Configuration -Configuration $script:Config -Path $ConfigPath
                    }
                }
            }
            
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "3" {
            # Remove remote target
            if ($script:RemoteTargets.Count -eq 0) {
                Write-Host "No targets to remove." -ForegroundColor $script:Colors.Warning
                Start-Sleep -Seconds 1
                Show-RemoteExecutionSettings
                return
            }
            
            Write-Host "Select a target to remove:" -ForegroundColor $script:Colors.Info
            
            for ($i = 0; $i -lt $script:RemoteTargets.Count; $i++) {
                Write-Host "$($i+1). $($script:RemoteTargets[$i])" -ForegroundColor $script:Colors.Success
            }
            
            $targetChoice = Show-Prompt -Message "Enter target number to remove (or C to cancel)"
            
            if ($targetChoice -eq "C") {
                Show-RemoteExecutionSettings
                return
            }
            
            $parsedChoice = $null
            if ([int]::TryParse($targetChoice, [ref]$parsedChoice) -and $parsedChoice -gt 0 -and $parsedChoice -le $script:RemoteTargets.Count) {
                $targetToRemove = $script:RemoteTargets[$parsedChoice - 1]
                $script:RemoteTargets = $script:RemoteTargets | Where-Object { $_ -ne $targetToRemove }
                $script:Config.RemoteExecution.DefaultTargets = $script:Config.RemoteExecution.DefaultTargets | Where-Object { $_ -ne $targetToRemove }
                
                Write-Host "Target $targetToRemove removed." -ForegroundColor $script:Colors.Success
                Export-Configuration -Configuration $script:Config -Path $ConfigPath
            }
            else {
                Write-Host "Invalid selection." -ForegroundColor $script:Colors.Error
            }
            
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "4" {
            # Test connectivity
            if ($script:RemoteTargets.Count -eq 0) {
                Write-Host "No targets to test." -ForegroundColor $script:Colors.Warning
                Start-Sleep -Seconds 1
                Show-RemoteExecutionSettings
                return
            }
            
            Write-Host "Testing connectivity to all targets..." -ForegroundColor $script:Colors.Info
            
            foreach ($target in $script:RemoteTargets) {
                $isReachable = Test-Connection -ComputerName $target -Count 1 -Quiet
                $statusColor = if ($isReachable) { $script:Colors.Success } else { $script:Colors.Error }
                $statusText = if ($isReachable) { "Reachable" } else { "Unreachable" }
                
                Write-Host "  $target - $statusText" -ForegroundColor $statusColor
                
                if ($isReachable) {
                    # Test PowerShell remoting
                    try {
                        Invoke-Command -ComputerName $target -ScriptBlock { "OK" } -ErrorAction Stop | Out-Null
                        Write-Host "    PowerShell Remoting: " -NoNewline -ForegroundColor $script:Colors.Info
                        Write-Host "Enabled" -ForegroundColor $script:Colors.Success
                    }
                    catch {
                        Write-Host "    PowerShell Remoting: " -NoNewline -ForegroundColor $script:Colors.Info
                        Write-Host "Disabled or Error" -ForegroundColor $script:Colors.Error
                        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor $script:Colors.Error
                    }
                }
            }
            
            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor $script:Colors.Warning
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            Show-RemoteExecutionSettings
        }
        "5" {
            # Change authentication method
            Write-Host "Select authentication method:" -ForegroundColor $script:Colors.Info
            Write-Host "1. Credential (username/password)" -ForegroundColor $script:Colors.Success
            Write-Host "2. Certificate" -ForegroundColor $script:Colors.Success
            Write-Host "3. CredSSP (enables credential delegation)" -ForegroundColor $script:Colors.Success
            
            $authChoice = Show-Prompt -Message "Enter choice" -Options @("1", "2", "3")
            
            switch ($authChoice) {
                "1" { $script:Config.RemoteExecution.AuthenticationMethod = "Credential" }
                "2" { $script:Config.RemoteExecution.AuthenticationMethod = "Certificate" }
                "3" { $script:Config.RemoteExecution.AuthenticationMethod = "CredSSP" }
            }
            
            Write-Host "Authentication method updated to $($script:Config.RemoteExecution.AuthenticationMethod)." -ForegroundColor $script:Colors.Success
            Export-Configuration -Configuration $script:Config -Path $ConfigPath
            
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "6" {
            # Update max concurrent sessions - FIXED VALIDATION LOGIC
            $maxConcurrent = Show-Prompt -Message "Enter maximum concurrent sessions (1-20)"
            
            if ($maxConcurrent -match '^\d+$' -and [int]$maxConcurrent -ge 1 -and [int]$maxConcurrent -le 20) {
                $script:Config.RemoteExecution.MaxConcurrentSessions = [int]$maxConcurrent
                Write-Host "Max concurrent sessions updated to $maxConcurrent." -ForegroundColor $script:Colors.Success
                Export-Configuration -Configuration $script:Config -Path $ConfigPath
            }
            else {
                Write-Host "Invalid value. Must be between 1 and 20." -ForegroundColor $script:Colors.Error
            }
            
            Start-Sleep -Seconds 1
            Show-RemoteExecutionSettings
        }
        "B" {
            Show-MainMenu
        }
    }
}

# Optimized Fix for Show-AIAssistantChat Function
function Show-AIAssistantChat {
    [CmdletBinding()]
    [OutputType([void])]
    param()

    Show-Header -Title "AI Troubleshooting Assistant"

    # Initialize UI with optimized console sizing
    $consoleWidth = Get-ConsoleWidth
    $chatBoxWidth = $consoleWidth - 4
    $maxDisplayedMessages = [Math]::Max(10, ($Host.UI.RawUI.WindowSize.Height - 15))

    Write-Host "Welcome to the AI Troubleshooting Assistant!" -ForegroundColor $script:Colors.Success
    Write-Host "Ask technical questions or describe issues you're experiencing." -ForegroundColor $script:Colors.Info
    Write-Host "Type 'exit', 'quit', or 'back' to return to the main menu." -ForegroundColor $script:Colors.Warning
    Write-Host ""

    # Session state initialization with proper type declaration
    $chatHistory = [System.Collections.Generic.List[hashtable]]::new()
    $exitChat = $false

    # Start UI refresh job with proper cleanup handling
    $refreshJob = $null
    if ($script:Config.UI.EnableAnimations -and $script:Config.UI.RefreshRate -gt 0) {
        $refreshJobScript = {
            param([int]$refreshRate)
            while ($true) {
                [Console]::Write("`n`n[Chat session active]`n")
                Start-Sleep -Seconds $refreshRate
            }
        }
        $refreshJob = Start-Job -ScriptBlock $refreshJobScript -ArgumentList $script:Config.UI.RefreshRate
    }

    try {
        while (-not $exitChat) {
            # Render chat history with optimized display logic
            if ($chatHistory.Count -gt 0) {
                $startIndex = [Math]::Max(0, $chatHistory.Count - $maxDisplayedMessages)
                $messagesToShow = $chatHistory.GetRange($startIndex, $chatHistory.Count - $startIndex)
                
                foreach ($message in $messagesToShow) {
                    $roleColor = if ($message.Role -eq "User") { $script:Colors.Primary } else { $script:Colors.Secondary }
                    $roleName = if ($message.Role -eq "User") { "You" } else { "Assistant" }

                    Write-Host "${roleName}:" -ForegroundColor $roleColor

                    # Format message text with efficient code block handling
                    $inCodeBlock = $false
                    foreach ($line in $message.Content -split "`n") {
                        if ($line -match '```(powershell|ps|bash|cmd|batch)?') {
                            $inCodeBlock = $true
                            Write-Host ""
                            continue
                        }
                        elseif ($line -match '```' -and $inCodeBlock) {
                            $inCodeBlock = $false
                            Write-Host ""
                            continue
                        }

                        if ($inCodeBlock) {
                            Write-Host $line -ForegroundColor $script:Colors.Accent
                        }
                        else {
                            # Efficient text wrapping with minimal string operations
                            if ($line.Length -gt $chatBoxWidth) {
                                for ($i = 0; $i -lt $line.Length; $i += $chatBoxWidth) {
                                    $wrappedLine = $line.Substring($i, [Math]::Min($chatBoxWidth, $line.Length - $i))
                                    Write-Host $wrappedLine -ForegroundColor $script:Colors.Info
                                }
                            }
                            else {
                                Write-Host $line -ForegroundColor $script:Colors.Info
                            }
                        }
                    }
                    Write-Host ""
                }
            }

            # Get user$inputFromUser with clear prompt
            Write-Host "You: " -NoNewline -ForegroundColor $script:Colors.Primary
            $userQuery = Read-Host

            # Check for exit commands efficiently
            if ($userQuery -in @('exit', 'quit', 'back')) {
                $exitChat = $true
                continue
            }

            # Add user message to history with correct timestamp
            $chatHistory.Add(@{
                Role = "User"
                Content = $userQuery
                Timestamp = Get-Date
            })

            # Show processing indicator
            Write-Host "`nProcessing..." -NoNewline -ForegroundColor $script:Colors.Info

            # Prepare AI prompt with optimized context building
            $contextPrompt = if ($chatHistory.Count -gt 1) {
                # Include limited context with efficient string building
                $contextSize = [Math]::Min(10, $chatHistory.Count - 1)
                $contextMessages = $chatHistory.GetRange(0, $chatHistory.Count - 1).GetRange(
                    [Math]::Max(0, $chatHistory.Count - 1 - $contextSize), 
                    [Math]::Min($contextSize, $chatHistory.Count - 1)
                )
                
                $sb = [System.Text.StringBuilder]::new("Previous conversation:`n")
                foreach ($msg in $contextMessages) {
                    [void]$sb.AppendLine("$($msg.Role): $($msg.Content)")
                }
                [void]$sb.AppendLine("`nNow answer this: $userQuery")
                $sb.ToString()
            }
            else {
                $userQuery
            }

            # Get response from AI with enhanced error handling
            try {
                # Check for automation opportunities
                $potentialActions = Find-AutomationOpportunities -Query $userQuery

                if ($potentialActions.Count -gt 0 -and $potentialActions[0].Confidence -gt 0.8) {
                    # Clear processing indicator
                    Write-Host "`r$(' ' * 20)`r" -NoNewline

                    # Suggest automation with clear options
                    $action = $potentialActions[0]
                    Write-Host "`nAssistant:" -ForegroundColor $script:Colors.Secondary
                    Write-Host "I can help with that. Would you like me to run the '$($action.ScriptName)' script to address this issue?" -ForegroundColor $script:Colors.Info
                    
                    $runScript = Show-Prompt -Message "Run script now? (Y/N)" -Options @("Y", "N")
                    
                    if ($runScript -eq "Y") {
                        $script = $script:ScriptDefinitions | Where-Object { $_.Description -eq $action.ScriptName } | Select-Object -First 1
                        
                        if ($script) {
                            $response = "Running the $($action.ScriptName) script...`n`nThis should help address your issue with $($action.Topic)."
                            
                            # Execute script with proper result handling
                            $success = Invoke-ScriptFromUrl -Url $script.ScriptUrl -ScriptName $script.Description
                            
                            $response += if ($success) {
                                "`n`nScript executed successfully. Is there anything else you'd like me to explain or help with?"
                            } else {
                                "`n`nThe script encountered an error during execution. Would you like me to suggest a different approach?"
                            }
                        }
                        else {
                            $response = "I recommended running a script that I couldn't find in the current script database. Let me provide some manual troubleshooting steps instead.`n`n"
                            $response += Invoke-AIEngine -Prompt $contextPrompt
                        }
                    }
                    else {
                        # User declined automation, get normal response
                        $response = Invoke-AIEngine -Prompt $contextPrompt
                    }
                }
                else {
                    # Standard AI response
                    $response = Invoke-AIEngine -Prompt $contextPrompt
                    
                    # Clear processing indicator
                    Write-Host "`r$(' ' * 20)`r" -NoNewline
                }
            }
            catch {
                # Improved error handling with detailed logging
                Write-Host "`r$(' ' * 20)`r" -NoNewline
                $errorDetails = "$($_.Exception.Message)`n$($_.Exception.StackTrace)"
                Write-Log "AI processing error: $errorDetails" -Level ERROR
                $response = "I'm sorry, I experienced an error while processing your request: $($_.Exception.Message)`n`nPlease try again with a different query or check the AI configuration."
            }
            
            # Add AI response to history
            $chatHistory.Add(@{
                Role = "Assistant"
                Content = $response
                Timestamp = Get-Date
            })
            
            # Display response with optimized formatting
            Write-Host "`nAssistant:" -ForegroundColor $script:Colors.Secondary
            
            # Format code blocks with efficient detection
            $inCodeBlock = $false
            foreach ($line in $response -split "`n") {
                if ($line -match '```(powershell|ps|bash|cmd|batch)?') {
                    $inCodeBlock = $true
                    Write-Host ""
                    continue
                }
                elseif ($line -match '```' -and $inCodeBlock) {
                    $inCodeBlock = $false
                    Write-Host ""
                    continue
                }
                
                # Apply appropriate formatting
                if ($inCodeBlock) {
                    Write-Host $line -ForegroundColor $script:Colors.Accent
                }
                else {
                    Write-Host $line -ForegroundColor $script:Colors.Info
                }
            }
            
            Write-Host ""
            
            # Log telemetry with enhanced metadata
            Send-Telemetry -EventName "AIChatMessage" -EventData @{
                QueryLength = $userQuery.Length
                ResponseLength = $response.Length
                ChatHistorySize = $chatHistory.Count
                TimestampUTC = [DateTime]::UtcNow.ToString('o')
            }
            
            # Check for script recommendations with improved pattern matching
            if ($response -match "I recommend running (\w+)" -or $response -match "try running the ([A-Za-z\s]+) script") {
                $recommendedScript = $matches[1]
                $matchingScripts = $script:ScriptDefinitions | 
                                   Where-Object { $_.Description -like "*$recommendedScript*" } |
                                   Select-Object -First 3
                
                if ($matchingScripts.Count -gt 0) {
                    Write-Host "`nI found these scripts that might help:" -ForegroundColor $script:Colors.Warning
                    
                    $index = 1
                    foreach ($script in $matchingScripts) {
                        Write-Host "$index. $($script.Description)" -ForegroundColor $script:Colors.Success
                        $index++
                    }
                    
                    Write-Host "`nWould you like to run one of these scripts? (Enter number or N)" -ForegroundColor $script:Colors.Warning
                    $scriptChoice = Read-Host
                    
                    # Fixed validation logic
                    if ($scriptChoice -match '^\d+$' -and [int]$scriptChoice -ge 1 -and [int]$scriptChoice -le $matchingScripts.Count) {
                        $selectedScript = $matchingScripts[[int]$scriptChoice - 1]
                        
                        # Execute the script
                        $success = Invoke-ScriptFromUrl -Url $selectedScript.ScriptUrl -ScriptName $selectedScript.Description
                        
                        Write-Host $(if ($success) {
                            "Script executed successfully."
                        } else {
                            "Script execution failed."
                        }) -ForegroundColor $(if ($success) { $script:Colors.Success } else { $script:Colors.Error })
                    }
                }
            }
        }
    }
    finally {
        # Proper cleanup of background job
        if ($null -ne $refreshJob) {
            Stop-Job -Job $refreshJob -ErrorAction SilentlyContinue
            Remove-Job -Job $refreshJob -Force -ErrorAction SilentlyContinue
        }
    }
    
    Show-MainMenu
}

# Optimized Find-AutomationOpportunities Function with Proper Regex Handling
function Find-AutomationOpportunities {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[hashtable]])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Query
    )
    
    # Use generic list for better performance with large collections
    $opportunities = [System.Collections.Generic.List[hashtable]]::new()
    
    # Define patterns with proper regex syntax
    $patterns = @(
        @{
            Pattern = 'clean.*?(disk|space|drive|storage|temp)'
            ScriptName = 'Disk Cleaner'
            Topic = 'disk space'
            Confidence = 0.85
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/DiskCleaner.ps1"
        },
        @{
            Pattern = '(remove|uninstall).*?bloatware'
            ScriptName = 'Dell Bloatware Remover'
            Topic = 'bloatware'
            Confidence = 0.9
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/BloatWareRemover.ps1"
        },
        @{
            Pattern = '(check|space|storage|disk).*?drive'
            ScriptName = 'Check Drive Space'
            Topic = 'drive space'
            Confidence = 0.8
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/CheckDriveSpace.ps1"
        },
        @{
            Pattern = 'install.*?(windows|microsoft).*?update'
            ScriptName = 'Reset & Install Windows Update'
            Topic = 'Windows updates'
            Confidence = 0.9
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/InstallWindowsUpdate.ps1"
        },
        @{
            Pattern = '(internet|network).*?(speed|throughput|bandwidth)'
            ScriptName = 'Internet Speed Test'
            Topic = 'internet speed'
            Confidence = 0.85
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/InternetSpeedTest.ps1"
        },
        @{
            Pattern = '(internet|network).*?(latency|ping|response)'
            ScriptName = 'Internet Latency Test'
            Topic = 'network latency'
            Confidence = 0.85
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/InternetLatencyTest.ps1"
        },
        @{
            Pattern = '(repair|fix).*?(windows|system)'
            ScriptName = 'Windows System Repair'
            Topic = 'system repair'
            Confidence = 0.8
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/WindowsSystemRepair.ps1"
        },
        @{
            Pattern = '(search|windows\s+search).*?not\s+working'
            ScriptName = 'Reset Windows Search DB'
            Topic = 'Windows Search'
            Confidence = 0.9
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/ResetandClearWindowsSearchDB.ps1"
        },
        @{
            Pattern = 'workpaper.*?monitor'
            ScriptName = 'WorkPaper Monitor Troubleshooter'
            Topic = 'WorkPaper Monitor'
            Confidence = 0.95
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/WorkPaperMonitorTroubleShooter.ps1"
        },
        @{
            Pattern = 'profile.*?(issue|problem|corrupt)'
            ScriptName = 'Check User Profile'
            Topic = 'user profile'
            Confidence = 0.85
            ScriptUrl = "$($script:Config.ScriptRepository.Uri)/CheckUserProfileIssue.ps1"
        }
    )
    
    # Optimized pattern matching with case-insensitive flag
    foreach ($pattern in $patterns) {
        try {
            if ($Query -match $pattern.Pattern) {
                $opportunities.Add($pattern)
            }
        }
        catch {
            Write-Log "Pattern matching error with pattern '$($pattern.Pattern)': $_" -Level ERROR
        }
    }
    
    # Sort results by confidence with LINQ-style approach for better performance
    return $opportunities | Sort-Object -Property Confidence -Descending
}

# Optimized Invoke-ScriptFromUrl Function (Error-Prone Section)
function Invoke-ScriptFromUrl {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Url,
        
        [Parameter()]
        [string]$ScriptName = (Split-Path -Path $Url -Leaf),
        
        [Parameter()]
        [hashtable]$Parameters = @{},
        
        [Parameter()]
        [string[]]$RemoteTargets,
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    # Measure execution time for performance metrics
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        Write-Log "Executing script from URL: $Url" -Level INFO
        
        # Early connectivity validation
        if (!(Test-InternetConnectivity)) {
            throw [System.Net.NetworkInformation.NetworkException]::new("No internet connection available")
        }
        
        # Create unique temporary file with security considerations
        $tempScriptPath = Join-Path -Path $script:TempDirectory -ChildPath "AutoByte_$([Guid]::NewGuid().ToString()).ps1"
        
        try {
            # Set TLS 1.2 for secure download
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            # Prepare web client with proper headers
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "AutoByte/$script:AppVersion PowerShell/$($PSVersionTable.PSVersion)")
            
            # Handle authentication if required
            if ($script:Config.ScriptRepository.AuthRequired) {
                $apiKey = Get-SecureApiKey -Prompt
                if ($null -ne $apiKey) {
                    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey)
                    $plainApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                    
                    $webClient.Headers.Add("Authorization", "Bearer $plainApiKey")
                }
                else {
                    throw [System.Security.Authentication.AuthenticationException]::new("API key required but not provided")
                }
            }
            
            # Download and verify script
            $scriptContent = $webClient.DownloadString($Url)
            
            if ($script:Config.Security.RequireSignedScripts -and -not (Test-ScriptSignature -ScriptContent $scriptContent)) {
                throw [System.Security.SecurityException]::new("Script signature validation failed. Execution aborted.")
            }
            
            # Save to temporary file with security considerations
            Set-Content -Path $tempScriptPath -Value $scriptContent -Force
            Write-Log "Script downloaded successfully to: $tempScriptPath" -Level DEBUG
        }
        catch [System.Net.WebException] {
            throw [System.Net.WebException]::new("Failed to download script: $($_.Exception.Message)", $_.Exception)
        }
        
        # Determine execution mode (remote vs local)
        if (($RemoteTargets -and $RemoteTargets.Count -gt 0) -or 
            ($script:Config.RemoteExecution.Enabled -and $script:RemoteTargets.Count -gt 0)) {
            
            $targetComputers = if ($RemoteTargets) { $RemoteTargets } else { $script:RemoteTargets }
            $remoteResult = Invoke-ScriptRemotely -ScriptPath $tempScriptPath -ScriptName $ScriptName -ComputerNames $targetComputers -Credential $Credential -Parameters $Parameters
            
            $stopwatch.Stop()
            $executionTime = $stopwatch.Elapsed
            
            Write-Log "Remote script execution completed in $($executionTime.TotalSeconds.ToString("0.00")) seconds" -Level SUCCESS
            
            # Log telemetry with enhanced metrics
            Send-Telemetry -EventName "ScriptExecuted" -EventData @{
                ScriptName = $ScriptName
                ExecutionTime = $executionTime.TotalSeconds
                Remote = $true
                TargetCount = $targetComputers.Count
                Success = $remoteResult.Success
                Timestamp = [DateTime]::UtcNow.ToString('o')
            }
            
            # Clean up temporary file immediately
            if (Test-Path -Path $tempScriptPath) {
                Remove-Item -Path $tempScriptPath -Force -ErrorAction SilentlyContinue
            }
            
            return $remoteResult.Success
        }
        else {
            # Local execution path
            Write-Log "Executing script: $ScriptName" -Level INFO
            Write-Host "`n" -NoNewline
            Write-Host "Executing $ScriptName..." -ForegroundColor $script:Colors.Primary
            Write-Host "----------------------------------------------------------------" -ForegroundColor $script:Colors.Primary
            
            # Efficient parameter string building
            $paramString = ""
            if ($Parameters.Count -gt 0) {
                $paramArray = foreach ($key in $Parameters.Keys) {
                    $value = $Parameters[$key]
                    if ($value -is [switch] -and $value) {
                        "-$key"
                    }
                    elseif ($value -is [string]) {
                        "-$key '$value'"
                    }
                    else {
                        "-$key $value"
                    }
                }
                $paramString = $paramArray -join " "
            }
            
            # Execution block with parameter handling
            $executionBlock = {
                param($path, $params)
                
                if ([string]::IsNullOrEmpty($params)) {
                    & $path
                }
                else {
                    # Using Invoke-Expression is controlled here since we build the parameters
                    $command = "& '$path' $params"
                    Invoke-Expression $command
                }
            }
            
            # Execute with optimized error handling
            $originalErrorAction = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            
            try {
                # Capture all output including errors
                $executionResult = & $executionBlock $tempScriptPath $paramString 2>&1
                
                # Separate output and errors for proper handling
                $outputStream = $executionResult | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }
                $errorOutput = $executionResult | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }
                
                # Display regular output
                $outputStream | ForEach-Object { Write-Host $_ }
                
                # Handle and display errors if any
                $success = $true
                if ($null -ne $errorOutput -and $errorOutput.Count -gt 0) {
                    Write-Host "`nErrors encountered during execution:" -ForegroundColor $script:Colors.Error
                    
                    foreach ($error in $errorOutput) {
                        Write-Host "- $($error.Exception.Message)" -ForegroundColor $script:Colors.Error
                        # Log errors for troubleshooting
                        Write-Log "Script execution error: $($error.Exception.Message)" -Level ERROR
                    }
                    
                    $success = $false
                }
            }
            catch {
                Write-Host "Error executing script: $($_.Exception.Message)" -ForegroundColor $script:Colors.Error
                Write-Log "Fatal error executing script: $_" -Level ERROR
                $success = $false
            }
            finally {
                # Restore original error preference
                $ErrorActionPreference = $originalErrorAction
            }
            
            Write-Host "----------------------------------------------------------------" -ForegroundColor $script:Colors.Primary
            
            $stopwatch.Stop()
            $executionTime = $stopwatch.Elapsed
            
            # Log execution result
            if ($success) {
                Write-Log "Script execution completed successfully in $($executionTime.TotalSeconds.ToString("0.00")) seconds: $ScriptName" -Level SUCCESS
            }
            else {
                Write-Log "Script execution failed in $($executionTime.TotalSeconds.ToString("0.00")) seconds: $ScriptName" -Level ERROR
            }
            
            # Log telemetry with comprehensive metrics
            Send-Telemetry -EventName "ScriptExecuted" -EventData @{
                ScriptName = $ScriptName
                ExecutionTime = $executionTime.TotalSeconds
                Remote = $false
                Success = $success
                ErrorCount = if ($errorOutput) { $errorOutput.Count } else { 0 }
                Timestamp = [DateTime]::UtcNow.ToString('o')
            }
            
            # Clean up temporary file with reliable removal
            if (Test-Path -Path $tempScriptPath) {
                Remove-Item -Path $tempScriptPath -Force -ErrorAction SilentlyContinue
            }
            
            return $success
        }
    }
    catch {
        # Handle execution failure gracefully
        $stopwatch.Stop()
        $executionTime = $stopwatch.Elapsed
        
        # Enhanced error logging with exception details
        $errorDetails = @{
            Message = $_.Exception.Message
            Type = $_.Exception.GetType().Name
            StackTrace = $_.Exception.StackTrace
            InnerException = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $null }
        }
        
        Write-Log "Error executing script from URL: $($errorDetails.Message)" -Level ERROR
        Write-Host "Error: $($errorDetails.Message)" -ForegroundColor $script:Colors.Error
        
        # Log detailed telemetry for troubleshooting
        Send-Telemetry -EventName "ScriptExecutionError" -EventData @{
            ScriptName = $ScriptName
            ExecutionTime = $executionTime.TotalSeconds
            ErrorMessage = $errorDetails.Message
            ErrorType = $errorDetails.Type
            Timestamp = [DateTime]::UtcNow.ToString('o')
        }
        
        # Clean up on failure with reliability
        if (Test-Path -Path $tempScriptPath) {
            Remove-Item -Path $tempScriptPath -Force -ErrorAction SilentlyContinue
        }
        
        return $false
    }
}

# Display available scripts
Write-Host "Select a script to execute:" -ForegroundColor White

# Display available scripts
for ($i = 0; $i -lt $script:ScriptDefinitions.Count; $i++) {
    Write-Host "$($i+1). $($script:ScriptDefinitions[$i].Description)" -ForegroundColor Green
}

function Show-Prompt {
    param (
        [string]$Message,
        [string[]]$Options
    )
    do {
        Write-Host $Message -ForegroundColor Cyan
        $inputFromUser = Read-Host
        if ($Options -contains $inputFromUser) {
            return $inputFromUser
        } else {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
        }
    } while ($true)
}

# Display available scripts
for ($i = 0; $i -lt $script:ScriptDefinitions.Count; $i++) {
    Write-Host "$($i+1). $($script:ScriptDefinitions[$i].Description)" -ForegroundColor Green
}

$scriptChoice = Show-Prompt -Message "Enter script number to execute (or C to cancel)" -Options @("C") + (1..$script:ScriptDefinitions.Count | ForEach-Object { "$_" })

if ($scriptChoice -eq "C") {
    Write-Host "Operation canceled." -ForegroundColor Red
    return
}

$parsedChoice = $null
if ([int]::TryParse($scriptChoice, [ref]$parsedChoice) -and $parsedChoice -gt 0 -and $parsedChoice -le $script:ScriptDefinitions.Count) {
    $selectedScript = $script:ScriptDefinitions[$parsedChoice - 1]
    Write-Host "Executing script: $($selectedScript.Description)" -ForegroundColor Cyan

    # Execute the selected script
    $success = Invoke-ScriptFromUrl -Url $selectedScript.ScriptUrl -ScriptName $selectedScript.Description

    if ($success) {
        Write-Host "Script executed successfully." -ForegroundColor Green
    } else {
        Write-Host "Script execution failed." -ForegroundColor Red
    }
} else {
    Write-Host "Invalid selection." -ForegroundColor Red
}

