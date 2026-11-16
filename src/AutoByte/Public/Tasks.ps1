function Get-AbTask {
    <#
    .SYNOPSIS
        Lists AutoByte tasks (checks and fixes) and their metadata.
    .EXAMPLE
        Get-AbTask -Kind Fix | Format-Table Id, Name, RequiresAdmin
    #>
    [CmdletBinding()]
    param([string[]]$Id, [ValidateSet('Check', 'Fix')][string]$Kind, [string]$Category)
    foreach ($t in $script:AbTasks.Values) {
        if ($Id -and $t.Id -notin $Id) { continue }
        if ($Kind -and $t.Kind -ne $Kind) { continue }
        if ($Category -and $t.Category -ne $Category) { continue }
        [pscustomobject]@{
            PSTypeName          = 'AutoByte.TaskInfo'
            Id                  = $t.Id
            Name                = $t.Name
            Kind                = $t.Kind
            Category            = $t.Category
            Description         = $t.Description
            Mutating            = [bool]$t.Mutating
            RequiresAdmin       = [bool]$t.RequiresAdmin
            RebootMayBeRequired = [bool]$t.RebootMayBeRequired
            EndUserSafe         = [bool]$t.EndUserSafe
            Parameters          = $t.Parameters
        }
    }
}

function Invoke-AbTask {
    <#
    .SYNOPSIS
        Runs one task and returns an AutoByte.Result.
    .DESCRIPTION
        Check tasks only read state and return findings.
        Fix tasks are idempotent: Test decides whether the fix is needed; if not, nothing changes
        (Status NotNeeded). Otherwise Repair runs (honours -WhatIf/-Confirm), then Test runs again
        to verify. A fix that still appears needed afterwards is reported Failed, unless it reported
        that a reboot is required to complete.
    .PARAMETER Parameters
        Overrides for the task's declared parameters, e.g. @{ MinAgeDays = 14 }.
    .EXAMPLE
        Invoke-AbTask -Id disk-space
    .EXAMPLE
        Invoke-AbTask -Id clean-temp-files -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)][string]$Id,
        [hashtable]$Parameters
    )
    process {
        $def = Get-AbTaskDefinition -Id $Id
        $result = New-AbResult -TaskId $def.Id -Kind $def.Kind
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $ctx = @{ Config = (Get-AbConfig); Result = $result; AfterRepair = $false }
        try {
            $p = Merge-AbParameter -Defaults $def.Parameters -Overrides $Parameters
            if ($def.RequiresAdmin -and -not (Test-AbIsAdmin)) {
                $result.Status = if ($def.Kind -eq 'Check') { 'Skipped' } else { 'Failed' }
                Add-AbMessage -Result $result -Message 'Requires an elevated (administrator) session.'
                return
            }
            if ($def.Kind -eq 'Check') {
                $findings = @(& $def.Test $ctx $p | Where-Object { $_ -and $_.PSObject.TypeNames -contains 'AutoByte.Finding' })
                $result.Findings = $findings
                $result.Status = if (@($findings | Where-Object Severity -NE 'Info').Count) { 'Issues' } else { 'Healthy' }
                return
            }

            $need = Get-AbVerdict (& $def.Test $ctx $p)
            if (-not $need.Needed) {
                $result.Status = 'NotNeeded'
                Add-AbMessage -Result $result -Message ("No action needed. " + $need.Reason).Trim()
                return
            }
            Add-AbMessage -Result $result -Message "Fix needed: $($need.Reason)"
            if (-not $PSCmdlet.ShouldProcess([Environment]::MachineName, $def.Name)) {
                $result.Status = 'WhatIf'
                Add-AbMessage -Result $result -Message "WhatIf: would run '$($def.Name)'."
                return
            }
            Write-AbLog -Message "[$($def.Id)] Repair starting."
            $out = @(& $def.Repair $ctx $p) | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
            if ($out -and $out['RebootRequired']) { $result.RebootRequired = $true }
            $result.Changed = $true
            $ctx.AfterRepair = $true   # lets Test ignore a -Force style parameter when verifying
            $after = Get-AbVerdict (& $def.Test $ctx $p)
            if ($after.Needed -and -not $result.RebootRequired) {
                $result.Status = 'Failed'
                Add-AbMessage -Result $result -Message "Verification failed: $($after.Reason)"
            }
            else {
                $result.Status = 'Fixed'
                if ($result.RebootRequired) { Add-AbMessage -Result $result -Message 'Fixed; a restart is required to complete.' }
                else { Add-AbMessage -Result $result -Message 'Fixed and verified.' }
            }
        }
        catch {
            $result.Status = 'Failed'
            $result.Error = $_.Exception.Message
            Write-AbLog -Level ERROR -Message "[$($def.Id)] $($_.Exception.Message)"
        }
        finally {
            $result.DurationMs = [int]$sw.Elapsed.TotalMilliseconds
            Write-AbLog -Message "[$($def.Id)] $($result.Status) in $($result.DurationMs) ms."
            $result
        }
    }
}

function Invoke-AbScan {
    <#
    .SYNOPSIS
        Runs every check task (read-only) and returns their results.
    .EXAMPLE
        $scan = Invoke-AbScan; $scan.Findings | Format-Table Severity, Title
    #>
    [CmdletBinding()]
    param([string[]]$Id, [string]$Category, [scriptblock]$OnProgress)
    $checks = @(Get-AbTask -Kind Check -Category $Category | Where-Object { -not $Id -or $_.Id -in $Id })
    $results = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($c in $checks) {
        $i++
        if ($OnProgress) { & $OnProgress $i $checks.Count $c.Name }
        $results.Add((Invoke-AbTask -Id $c.Id))
    }
    $all = @($results | ForEach-Object { $_.Findings })
    [pscustomobject]@{
        PSTypeName  = 'AutoByte.Scan'
        Computer    = [Environment]::MachineName
        StartedUtc  = if ($results.Count) { $results[0].StartedUtc } else { (Get-Date).ToUniversalTime().ToString('o') }
        Summary     = Get-AbSystemSummary
        Results     = $results.ToArray()
        Findings    = $all
        Critical    = @($all | Where-Object Severity -EQ 'Critical').Count
        Warning     = @($all | Where-Object Severity -EQ 'Warning').Count
        Info        = @($all | Where-Object Severity -EQ 'Info').Count
        SuggestedFixes = @($all | Where-Object { $_.FixTask -and $_.Severity -ne 'Info' } | ForEach-Object FixTask | Sort-Object -Unique)
    }
}

function Invoke-AbFix {
    <#
    .SYNOPSIS
        Runs fix tasks: the ones you name, or the ones a scan suggested.
    .PARAMETER Scan
        A result from Invoke-AbScan; runs its SuggestedFixes.
    .PARAMETER EndUserSafeOnly
        Only run fixes marked safe for non-technical users (used by the end-user menu).
    .EXAMPLE
        Invoke-AbScan | Invoke-AbFix -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(ValueFromPipeline)]$Scan,
        [string[]]$Id,
        [switch]$EndUserSafeOnly
    )
    process {
        $ids = @()
        if ($Id) { $ids += $Id }
        if ($Scan) { $ids += @($Scan.SuggestedFixes) }
        foreach ($t in ($ids | Sort-Object -Unique)) {
            $info = Get-AbTask -Id $t
            if (-not $info) { throw "Unknown task '$t'." }
            if ($info.Kind -ne 'Fix') { throw "'$t' is a check, not a fix." }
            if ($EndUserSafeOnly -and -not $info.EndUserSafe) {
                $r = New-AbResult -TaskId $t -Kind Fix
                $r.Status = 'Skipped'
                Add-AbMessage -Result $r -Message 'Needs a technician (not marked safe for end users).'
                $r
                continue
            }
            Invoke-AbTask -Id $t -WhatIf:$WhatIfPreference -Confirm:$false
        }
    }
}

function Get-AbSystemSummary {
    [CmdletBinding()]
    param()
    $s = [ordered]@{ Computer = [Environment]::MachineName; User = [Environment]::UserName; OS = ''; Version = ''; Model = ''; Manufacturer = ''; RamGB = 0; LastBoot = ''; AutoByte = $script:AbVersion }
    try {
        $os = Get-AbCim -ClassName Win32_OperatingSystem
        $cs = Get-AbCim -ClassName Win32_ComputerSystem
        $s.OS = "$($os.Caption)".Trim()
        $s.Version = "$($os.Version)"
        $s.LastBoot = if ($os.LastBootUpTime) { ([datetime]$os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm') } else { '' }
        $s.Manufacturer = "$($cs.Manufacturer)"
        $s.Model = "$($cs.Model)"
        $s.RamGB = [math]::Round([double]$cs.TotalPhysicalMemory / 1GB, 1)
    }
    catch { Write-AbLog -Level WARN -Message "System summary incomplete: $($_.Exception.Message)" }
    [pscustomobject]$s
}
