@{
    RootModule           = 'AutoByte.psm1'
    ModuleVersion        = '3.0.0'
    GUID                 = '6c0f2c55-5b0e-4c43-9a3f-1d6c3b8f7a21'
    Author               = 'deep1ne8'
    CompanyName          = 'AutoByte'
    Copyright            = '(c) AutoByte contributors'
    Description          = 'Windows health checks, safe idempotent fixes and HTML reports for MSP technicians, RMM (Datto) and end users.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @('Get-AbTask', 'Invoke-AbTask', 'Invoke-AbScan', 'Invoke-AbFix', 'Get-AbSystemSummary', 'Export-AbReport', 'ConvertTo-AbJson', 'Invoke-AbLegacyScript', 'Get-AbExitCode')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('Windows', 'MSP', 'RMM', 'Datto', 'Troubleshooting', 'Health')
            ProjectUri = 'https://github.com/deep1ne8/misc'
        }
    }
}
