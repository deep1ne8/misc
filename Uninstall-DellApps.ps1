# Define application details
$applications = @(
    @{
        Name = "Dell SupportAssist"
        Version = ""
    },
    @{
        Name = "Dell SupportAssist OS Recovery Plugin for Dell Update"
        Version = ""
    },
    @{
        Name = "Dell SupportAssist Remediation"
        Version = ""
    }
)

# Function to uninstall the application
function Uninstall-App {
    param (
        [string]$AppName,
        [string]$AppVersion
    )
    Write-Host "Attempting to uninstall $AppName..."
    $app = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -eq $AppName -and ($AppVersion -eq "" -or $_.Version -eq $AppVersion)
    }

    if ($app) {
        Write-Host "Uninstalling $AppName..."
        try {
            $app.Uninstall() | Out-Null
            Write-Host "$AppName uninstalled successfully."
        } catch {
            Write-Host "Failed to uninstall ${AppName}: $_"
        }
    } else {
        Write-Host "$AppName not found or already uninstalled."
    }
}

# Execution starts here
foreach ($application in $applications) {
    Uninstall-App -AppName $application.Name -AppVersion $application.Version
}
