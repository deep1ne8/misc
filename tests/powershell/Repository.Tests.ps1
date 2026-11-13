#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }
<#
.SYNOPSIS
    Repository-wide Pester 5 tests: parse, 5.1 syntax compatibility, standards ratchet,
    secret patterns, manifest freshness, and launcher behaviour.
#>

BeforeDiscovery {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
    $script:PsFiles = Get-ChildItem -Path $Root -Recurse -Force -Include *.ps1, *.psm1 -File |
        Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' -and $_.Length -gt 0 } |
        ForEach-Object { @{ Path = $_.FullName; Rel = $_.FullName.Substring($Root.Length + 1).Replace('\', '/') } }
    $baseline = Get-Content (Join-Path $PSScriptRoot 'legacy-baseline.txt') | Where-Object { $_.Trim() }
    $script:StandardFiles = $PsFiles | Where-Object { $_.Rel -like 'Scripts/*' -and $_.Rel -notin $baseline }
}

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..' '..')).Path
}

Describe 'Every PowerShell file' {
    It '<Rel> parses without errors' -ForEach $PsFiles {
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It '<Rel> uses no PowerShell 7-only syntax (5.1 endpoints)' -ForEach $PsFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        if ($text -match '(?im)^#Requires\s+-Version\s+7') { Set-ItResult -Skipped -Because 'declares #Requires -Version 7'; return }
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $ps7 = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.TernaryExpressionAst] -or
                $n -is [System.Management.Automation.Language.PipelineChainAst] -or
                ($n -is [System.Management.Automation.Language.BinaryExpressionAst] -and $n.Operator -in 'QuestionQuestion') -or
                ($n -is [System.Management.Automation.Language.AssignmentStatementAst] -and "$($n.Operator)" -eq 'QuestionQuestionEquals')
            }, $true)
        @($ps7 | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Extent.Text)" }) | Should -BeNullOrEmpty
    }

    It '<Rel> has a UTF-8 BOM if it contains non-ASCII (Windows PowerShell 5.1 misreads BOM-less UTF-8)' -ForEach $PsFiles {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $nonAscii = @($bytes | Where-Object { $_ -gt 127 }).Count -gt 0
        if ($nonAscii) { $hasBom | Should -BeTrue }
    }

    It '<Rel> contains no obvious embedded credentials' -ForEach $PsFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Not -Match '(?i)\$\w*(secret|password|passwd|apikey|accesskey)\w*\s*=\s*["''][A-Za-z0-9+/=_\-]{12,}["'']'
        $text | Should -Not -Match '(?i)\bKEY=[A-Za-z0-9+/]{30,}={0,2}'
    }
}

Describe 'Scripts outside the legacy baseline meet the header standard' {
    It '<Rel> has comment-based help and CmdletBinding' -ForEach $StandardFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Match '\.SYNOPSIS'
        $text | Should -Match 'CmdletBinding'
    }
}

Describe 'scripts.manifest.json' {
    It 'is up to date' {
        $out = & pwsh -NoProfile -File (Join-Path $Root 'build/Build-ScriptManifest.ps1') -Root $Root -Check 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($out | Out-String)
    }

    It 'never offers TPM/CPU-bypass scripts to the agent' {
        $m = Get-Content (Join-Path $Root 'scripts.manifest.json') -Raw | ConvertFrom-Json
        foreach ($s in $m.scripts | Where-Object { $_.name -match '(?i)bypass|tpm' }) {
            $s.unsupported | Should -BeTrue
            $s.agentDiagnostic | Should -BeFalse
        }
    }

    It 'only marks non-mutating, non-interactive scripts as agent diagnostics' {
        $m = Get-Content (Join-Path $Root 'scripts.manifest.json') -Raw | ConvertFrom-Json
        foreach ($s in $m.scripts | Where-Object agentDiagnostic) {
            $s.mutating | Should -BeFalse
            $s.interactive | Should -BeFalse
        }
    }
}

Describe 'AutoByte.ps1 launcher' {
    BeforeAll {
        . (Join-Path $Root 'AutoByte.ps1')
        $script:Manifest = Get-Content (Join-Path $Root 'scripts.manifest.json') -Raw | ConvertFrom-Json
    }

    It 'does not use Invoke-Expression' {
        (Get-Content (Join-Path $Root 'AutoByte.ps1') -Raw) | Should -Not -Match '(?i)\b(Invoke-Expression|iex)\b\s'
    }

    It 'defaults to a pinned ref' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $Root 'AutoByte.ps1'), [ref]$null, [ref]$null)
        $ref = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Ref' }
        $ref.DefaultValue.Value | Should -Match '^v\d+\.\d+\.\d+$'
    }

    It 'lists only menu scripts that exist in the manifest' {
        $MenuItems.Name | Where-Object { $_ -notin $Manifest.scripts.name } | Should -BeNullOrEmpty
    }

    It 'computes the same normalised hash as the manifest' {
        $entry = $Manifest.scripts | Where-Object name -EQ 'CheckDriveSpace'
        Get-NormalizedSha256 -Path (Join-Path $Root $entry.path) | Should -Be $entry.sha256
    }

    It 'treats CRLF and LF content identically' {
        $a = Join-Path $TestDrive 'a.ps1'; $b = Join-Path $TestDrive 'b.ps1'
        [IO.File]::WriteAllText($a, "x`r`ny`r`n"); [IO.File]::WriteAllText($b, "x`ny`n")
        Get-NormalizedSha256 -Path $a | Should -Be (Get-NormalizedSha256 -Path $b)
    }

    Context 'Invoke-VerifiedScript' {
        BeforeAll {
            $script:Ref = 'v2.0.0'; $script:Repository = 'deep1ne8/misc'
            $script:BaseUri = 'https://raw.githubusercontent.com'; $script:AllowUnsupported = $false
            $body = "Write-Output 'ok'`nexit 3010`n"
            $script:GoodHash = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-', '').ToLowerInvariant()
            $script:Fake = @{
                good   = [pscustomobject]@{ name = 'Good'; path = 'Scripts/Good.ps1'; sha256 = $GoodHash; platform = 'windows'; unsupported = $false }
                bypass = [pscustomobject]@{ name = 'Bypass'; path = 'Scripts/Bypass.ps1'; sha256 = $GoodHash; platform = 'windows'; unsupported = $true }
            }
            $script:Body = $body
        }

        It 'runs a verified script and returns its exit code' {
            Mock Invoke-WebRequest { [IO.File]::WriteAllText($OutFile, "Write-Output 'ok'`nexit 3010`n") }
            Invoke-VerifiedScript -Manifest $Fake -Name Good -Confirm:$false 6>$null | Should -Be 3010
        }

        It 'refuses a tampered download and does not run it' {
            Mock Invoke-WebRequest { [IO.File]::WriteAllText($OutFile, "Write-Output 'EVIL'; exit 0`n") }
            { Invoke-VerifiedScript -Manifest $Fake -Name Good -Confirm:$false 6>$null } | Should -Throw '*SHA-256 mismatch*'
        }

        It 'refuses unsupported scripts without -AllowUnsupported' {
            Mock Invoke-WebRequest { throw 'should not download' }
            { Invoke-VerifiedScript -Manifest $Fake -Name Bypass -Confirm:$false } | Should -Throw '*unsupported*'
        }

        It 'refuses names not in the manifest' {
            { Invoke-VerifiedScript -Manifest $Fake -Name Nope -Confirm:$false } | Should -Throw '*not in the manifest*'
        }

        It 'verifies but does not run under -WhatIf' {
            Mock Invoke-WebRequest { [IO.File]::WriteAllText($OutFile, "Write-Output 'ok'`nexit 3010`n") }
            Invoke-VerifiedScript -Manifest $Fake -Name Good -WhatIf 6>$null | Should -Be 0
        }
    }
}
