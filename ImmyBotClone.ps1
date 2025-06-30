#requires -Version 5.1
<#$
.SYNOPSIS
    ImmyBotClone - Basic automation framework similar to ImmyBot.

.DESCRIPTION
    Provides a simple device inventory and ability to run automation scripts
    on remote computers using PowerShell remoting. This is a lightweight
    demonstration and not a full replacement for ImmyBot.

.NOTES
    Author: Codex
    Version: 1.0
    Created: 2024-06-10
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$InventoryPath = "$PSScriptRoot\Devices.json",

    [Parameter()]
    [string]$ScriptRepository = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts"
)

function Initialize-Inventory {
    if (-not (Test-Path $InventoryPath)) {
        @() | ConvertTo-Json | Set-Content -Path $InventoryPath -Encoding UTF8
    }
}

function Get-Inventory {
    Initialize-Inventory
    return Get-Content -Path $InventoryPath | ConvertFrom-Json
}

function Save-Inventory ($inventory) {
    $inventory | ConvertTo-Json | Set-Content -Path $InventoryPath -Encoding UTF8
}

function Add-Device {
    param(
        [string]$ComputerName,
        [string]$Description
    )
    $inventory = Get-Inventory
    $inventory += [pscustomobject]@{
        ComputerName = $ComputerName
        Description  = $Description
    }
    Save-Inventory $inventory
    Write-Host "Device '$ComputerName' added." -ForegroundColor Green
}

function Remove-Device {
    param([string]$ComputerName)
    $inventory = Get-Inventory
    $newInventory = $inventory | Where-Object { $_.ComputerName -ne $ComputerName }
    Save-Inventory $newInventory
    Write-Host "Device '$ComputerName' removed." -ForegroundColor Yellow
}

function List-Devices {
    $inventory = Get-Inventory
    if ($inventory.Count -eq 0) {
        Write-Host "No devices found" -ForegroundColor Yellow
    } else {
        $inventory | Format-Table -AutoSize
    }
}

function Invoke-RemoteScript {
    param(
        [string]$ComputerName,
        [string]$ScriptName
    )
    $url = "$ScriptRepository/$ScriptName"
    try {
        $scriptContent = Invoke-WebRequest -Uri $url | Select-Object -ExpandProperty Content
        Invoke-Command -ComputerName $ComputerName -ScriptBlock ([ScriptBlock]::Create($scriptContent))
    } catch {
        Write-Host "Failed to execute script on $ComputerName" -ForegroundColor Red
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "ImmyBotClone" -ForegroundColor Cyan
    Write-Host "1. List devices"
    Write-Host "2. Add device"
    Write-Host "3. Remove device"
    Write-Host "4. Execute script on device"
    Write-Host "5. Exit"

    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { List-Devices }
        '2' {
            $name = Read-Host "Computer name"
            $desc = Read-Host "Description"
            Add-Device -ComputerName $name -Description $desc
        }
        '3' {
            $name = Read-Host "Computer name"
            Remove-Device -ComputerName $name
        }
        '4' {
            $name = Read-Host "Computer name"
            $script = Read-Host "Script name (e.g. DiskCleaner.ps1)"
            Invoke-RemoteScript -ComputerName $name -ScriptName $script
        }
        '5' { return }
        default { Write-Host "Invalid selection" -ForegroundColor Red }
    }
    Pause
    Show-Menu
}

Show-Menu
