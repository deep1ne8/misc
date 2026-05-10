# UnblockFiles.ps1
<#
.SYNOPSIS
    Recursively unblocks files downloaded from the internet (removes Zone.Identifier).

.PARAMETER FolderPath
    The path to the folder containing the files to unblock.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$FolderPath
)

if (-not (Test-Path $FolderPath)) {
    Write-Error "The folder path '$FolderPath' does not exist."
    exit 1
}

Write-Host "Unblocking files in: $FolderPath" -ForegroundColor Cyan

Get-ChildItem -Path $FolderPath -Recurse -File | ForEach-Object {
    try {
        Unblock-File -Path $_.FullName
        Write-Host "Unblocked: $($_.FullName)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to unblock: $($_.FullName). Error: $_"
    }
}

Write-Host "Completed unblocking files." -ForegroundColor Yellow
