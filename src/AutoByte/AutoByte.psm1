#Requires -Version 5.1
Set-StrictMode -Version Latest

$script:AbVersion = (Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'AutoByte.psd1')).ModuleVersion

foreach ($folder in 'Private', 'Public') {
    foreach ($file in Get-ChildItem -Path (Join-Path $PSScriptRoot $folder) -Filter *.ps1 -File | Sort-Object Name) {
        . $file.FullName
    }
}

# Each task file returns one hashtable. Dot-sourcing keeps its scriptblocks bound to this module,
# so they can call private helpers and adapters (and Pester can mock those adapters).
foreach ($file in Get-ChildItem -Path (Join-Path $PSScriptRoot 'Tasks') -Filter *.task.ps1 -File -Recurse | Sort-Object FullName) {
    $definition = . $file.FullName
    Register-AbTask -Definition $definition -Source $file.Name
}

Export-ModuleMember -Function Get-AbTask, Invoke-AbTask, Invoke-AbScan, Invoke-AbFix, Get-AbSystemSummary, Export-AbReport, ConvertTo-AbJson, Invoke-AbLegacyScript, Get-AbExitCode
