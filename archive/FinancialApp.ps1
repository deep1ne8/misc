# Financial AI Coach Web App Deployment Script
# This script automates the deployment of the Financial AI Coach web application
# to multiple machines or servers with minimal configuration.

param (
    [string]$DeploymentPath = "$env:USERPROFILE\FinanceAICoach",
    [string]$BackupPath = "$env:USERPROFILE\FinanceAICoach_Backup",
    [switch]$CreateBackup = $false,
    [switch]$RestoreFromBackup = $false,
    [switch]$InstallDependencies = $false,
    [string[]]$RemoteComputers = @()
)

# Logging function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    
    # Save to log file if needed
    $logFile = Join-Path -Path $DeploymentPath -ChildPath "deployment.log"
    Add-Content -Path $logFile -Value $logEntry
}

# Create app directory structure
function Initialize-AppStructure {
    Write-Log "Creating application directory structure..."
    
    if (-not (Test-Path -Path $DeploymentPath)) {
        New-Item -Path $DeploymentPath -ItemType Directory -Force | Out-Null
        Write-Log "Created main deployment directory: $DeploymentPath"
    }
    
    # Create index.html file with the web app content
    $indexPath = Join-Path -Path $DeploymentPath -ChildPath "index.html"
    
    # HTML content from artifact would be copied here
    $htmlContent = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Financial AI Coach</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/tailwindcss/2.2.19/tailwind.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        .gradient-bg {
            background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%);
        }
        .card {
            transition: all 0.3s ease;
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 10px 20px rgba(0, 0, 0, 0.1);
        }
        .progress-bar {
            height: 8px;
            border-radius: 4px;
            background: #e2e8f0;
            overflow: hidden;
        }
        .progress {
            height: 100%;
            border-radius: 4px;
            background: linear-gradient(90deg, #6366F1 0%, #8B5CF6 100%);
        }
    </style>
</head>
<body class="bg-gray-50 font-sans">
    <!-- Content would be here -->
</body>
</html>
'@

    Set-Content -Path $indexPath -Value $htmlContent
    Write-Log "Created index.html file at: $indexPath"
    
    # Create placeholder for user data
    $dataPath = Join-Path -Path $DeploymentPath -ChildPath "data"
    if (-not (Test-Path -Path $dataPath)) {
        New-Item -Path $dataPath -ItemType Directory -Force | Out-Null
        Write-Log "Created data directory: $dataPath"
        
        # Create empty JSON file for sample data
        $userDataPath = Join-Path -Path $dataPath -ChildPath "userData.json"
        Set-Content -Path $userDataPath -Value "{ `"users`": [] }"
        Write-Log "Created sample user data file"
    }
}

# Backup existing deployment
function Backup-Deployment {
    if ($CreateBackup) {
        Write-Log "Creating backup of existing deployment..."
        
        if (Test-Path -Path $DeploymentPath) {
            # Create backup directory if it doesn't exist
            if (-not (Test-Path -Path $BackupPath)) {
                New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
            }
            
            # Create timestamped backup
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $backupDir = Join-Path -Path $BackupPath -ChildPath "backup_$timestamp"
            
            # Copy contents
            Copy-Item -Path "$DeploymentPath\*" -Destination $backupDir -Recurse -Force
            Write-Log "Backup created at: $backupDir" 
        }
        else {
            Write-Log "No existing deployment found to backup" -Level "WARNING"
        }
    }
}

# Restore from backup
function Restore-Deployment {
    if ($RestoreFromBackup) {
        Write-Log "Attempting to restore from backup..."
        
        if (Test-Path -Path $BackupPath) {
            # Get latest backup
            $latestBackup = Get-ChildItem -Path $BackupPath -Directory | 
                            Sort-Object -Property LastWriteTime -Descending | 
                            Select-Object -First 1
            
            if ($latestBackup) {
                # Clear current deployment
                if (Test-Path -Path $DeploymentPath) {
                    Remove-Item -Path "$DeploymentPath\*" -Recurse -Force
                }
                else {
                    New-Item -Path $DeploymentPath -ItemType Directory -Force | Out-Null
                }
                
                # Copy from backup
                Copy-Item -Path "$($latestBackup.FullName)\*" -Destination $DeploymentPath -Recurse -Force
                Write-Log "Deployment restored from: $($latestBackup.FullName)"
            }
            else {
                Write-Log "No backups found to restore from" -Level "ERROR"
            }
        }
        else {
            Write-Log "Backup path does not exist: $BackupPath" -Level "ERROR"
        }
    }
}

# Install lightweight HTTP server if requested
function Install-Dependencies {
    if ($InstallDependencies) {
        Write-Log "Installing dependencies..."
        
        # Check if Node.js is installed
        $nodeInstalled = $null -ne (Get-Command "node" -ErrorAction SilentlyContinue)
        
        if (-not $nodeInstalled) {
            Write-Log "Node.js not found. Installing..." -Level "WARNING"
            
            # Download and install Node.js using winget if available
            $chocoInstalled = $null -ne (Get-Command "choco" -ErrorAction SilentlyContinue)
            
            if ($chocoInstalled) {
                Write-Log "Installing Node.js using choco..."
                choco install NodeJS -y
            }
            else {
                Write-Log "Choco not found. Please install Node.js manually." -Level "ERROR"
                return
            }
        }
    }
}

function SimpleHttpServer {
    # Define the port to listen on
    $port = 8080

    # Create a listener
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://*:8080/")

    $listener.Start()
    Write-Output "HTTP server started. Listening on port $port..."

    # Handle incoming requests
    # Use a while loop to keep listening until the server is stopped
    # This is a simple implementation and may not scale well for a production-grade application
    while ($listener.IsListening) {
    $context = $listener.GetContext()
    $response.StatusCode = 200
    $response.ContentType = "text/html"

    # $request = $context.Request
    $response = $context.Response
    $response.Headers.Add("Access-Control-Allow-Origin", "*")
    $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    # Read the index.html content from the deployment path
    $indexPath = Join-Path -Path $DeploymentPath -ChildPath "index.html"
    if (Test-Path -Path $indexPath) {
        $responseString = Get-Content -Path $indexPath -Raw
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
    } else {
        $responseString = "<html><body>index.html not found.</body></html>"
    }


    # Define the response content
    #$responseString = "<html><body>Hello, World!</body></html>"
    #$buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
    #$response.ContentLength64 = $buffer.Length
    #$response.OutputStream.Write($buffer, 0, $buffer.Length)
    #$response.OutputStream.Close()
    }

    # Stop the listener when done
    #$listener.Stop()
    #$listener.Close()
}




# Start the application locally
function Start-Application {
    Write-Log "Starting Financial AI Coach application..."
    
    $httpServerInstalled = $null -ne (Get-Command "http-server" -ErrorAction SilentlyContinue)
    
    if ($httpServerInstalled) {
        # Start http-server in the deployment directory
        #$serverProcess = Start-Process -FilePath "http-server" -ArgumentList $DeploymentPath -PassThru
        Write-Log "Application started. Server PID: $($serverProcess.Id)"
        Write-Log "Access the application at: http://localhost:8080"
    }
    else {
        Write-Log "http-server not found. Please install dependencies first." -Level "WARNING"
        Write-Log "You can still access the application by opening index.html directly in a browser."
    }
}


# Deploy to remote computers if specified
function Deploy-ToRemoteComputers {
    if ($RemoteComputers.Count -gt 0) {
        Write-Log "Deploying to remote computers: $($RemoteComputers -join ', ')"
        
        foreach ($computer in $RemoteComputers) {
            Write-Log "Deploying to $computer..."
            
            # Test connection
            if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
                # Create remote destination if it doesn't exist
                $remotePath = "\\$computer\C$\FinancialAICoach"
                
                # Create remote directory if doesn't exist
                if (-not (Test-Path -Path $remotePath)) {
                    New-Item -Path $remotePath -ItemType Directory -Force | Out-Null
                }
                
                # Copy files
                Copy-Item -Path "$DeploymentPath\*" -Destination $remotePath -Recurse -Force
                Write-Log "Deployment to $computer completed successfully"
            }
            else {
                Write-Log "Could not connect to $computer. Deployment failed." -Level "ERROR"
            }
        }
    }
}

# Main deployment process
try {
    Write-Log "Starting Financial AI Coach deployment process..."
    
    # Backup existing deployment if requested
    Backup-Deployment
    
    # Restore from backup if requested
    if ($RestoreFromBackup) {
        Restore-Deployment
    }
    else {
        # Create app structure for fresh install
        Initialize-AppStructure
        
        # Install dependencies if requested
        Install-Dependencies
        
        # Deploy to remote computers if specified
        Deploy-ToRemoteComputers
    }
    
    # Start the application locally
    SimpleHttpServer
    Start-Application
    
    Write-Log "Deployment process completed successfully."
}
catch {
    Write-Log "An error occurred during deployment: $_" -Level "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
}