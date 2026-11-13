@{
    # Full default rule set plus the repository's security/compatibility rules.
    IncludeDefaultRules = $true
    Severity            = @('Error', 'Warning')
    ExcludeRules        = @(
        # UI scripts legitimately use Write-Host for coloured console output.
        'PSAvoidUsingWriteHost'
    )
    Rules               = @{
        PSAvoidUsingInvokeExpression                = @{ Enable = $true }
        PSAvoidUsingPlainTextForPassword            = @{ Enable = $true }
        PSAvoidUsingConvertToSecureStringWithPlainText = @{ Enable = $true }
        PSUseShouldProcessForStateChangingFunctions = @{ Enable = $true }
        PSUseCompatibleSyntax                       = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.4')
        }
    }
}
