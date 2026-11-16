# CVE exposure: compares this PC's exact Windows build with the fixed builds Microsoft publishes
# per CVE (MSRC CVRF feed, keyless), and flags CVEs on CISA's Known Exploited Vulnerabilities list.
# Patches are applied through Windows Update (install-security-updates), never downloaded ad hoc.

function Get-AbMsrcProductName {
    # Candidate MSRC product names for this OS (MSRC naming varies slightly across releases).
    param([Parameter(Mandatory)]$Os)
    $arch = switch -Regex ($Os.Architecture) { 'ARM64' { 'ARM64-based Systems'; break } '^x86$' { '32-bit Systems'; break } default { 'x64-based Systems' } }
    $build = [int]$Os.Build
    $ver = [string]$Os.DisplayVersion
    if ($Os.InstallationType -match 'Server') {
        $name = switch ($build) {
            14393 { 'Windows Server 2016' } 17763 { 'Windows Server 2019' } 20348 { 'Windows Server 2022' }
            25398 { 'Windows Server 2022, 23H2 Edition' } 26100 { 'Windows Server 2025' } default { '' }
        }
        if (-not $name) { return @() }
        if ($Os.InstallationType -match 'Core') { return @("$name (Server Core installation)") }
        return @($name)
    }
    $family = if ($build -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    $names = @("$family Version $ver for $arch")
    if ($family -eq 'Windows 11' -and $ver -eq '21H2') { $names += "Windows 11 for $arch" }
    $names
}

function Get-AbCvrfProductId {
    # Walks a CVRF ProductTree (dictionaries or PSObjects) and returns ProductIDs whose name matches.
    param($Node, [string[]]$Names)
    if ($null -eq $Node -or $Node -is [string]) { return }
    if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [System.Collections.IDictionary])) {
        foreach ($n in $Node) { Get-AbCvrfProductId -Node $n -Names $Names }
        return
    }
    $id = Get-AbField $Node 'ProductID'; $val = Get-AbField $Node 'Value'
    if ($id -and $val -is [string] -and $Names -contains $val) { [string]$id }
    foreach ($k in 'FullProductName', 'Branch', 'Items') { $c = Get-AbField $Node $k; if ($c) { Get-AbCvrfProductId -Node $c -Names $Names } }
}

function Get-AbCveExposure {
    <#
      Pure function over parsed CVRF documents. Returns one object per CVE that is fixed in a build
      newer than CurrentBuild for one of ProductNames.
    #>
    param([object[]]$Documents, [Parameter(Mandatory)][string[]]$ProductNames, [Parameter(Mandatory)][version]$CurrentBuild, [string[]]$KevCves = @())
    $kev = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($k in $KevCves) { [void]$kev.Add($k) }
    $seen = @{}
    foreach ($doc in $Documents) {
        $pids = @(Get-AbCvrfProductId -Node (Get-AbField $doc 'ProductTree') -Names $ProductNames | Sort-Object -Unique)
        if (-not $pids) { continue }
        foreach ($v in @(Get-AbField $doc 'Vulnerability')) {
            $cve = [string](Get-AbField $v 'CVE')
            if (-not $cve -or $seen.ContainsKey($cve)) { continue }
            $best = $null
            foreach ($rem in @(Get-AbField $v 'Remediations')) {
                if ([string](Get-AbField $rem 'Type') -notin '2', 'Vendor Fix') { continue }
                $fb = [string](Get-AbField $rem 'FixedBuild')
                $remPids = @(Get-AbField $rem 'ProductID' | ForEach-Object { [string]$_ })
                if (-not $fb -or -not @($remPids | Where-Object { $pids -contains $_ })) { continue }
                $fixed = $null
                if (-not [version]::TryParse(($fb -replace '[^\d.]', ''), [ref]$fixed)) { continue }
                if ($fixed.Major -ne $CurrentBuild.Major -or $fixed.Build -ne $CurrentBuild.Build) { continue }  # other OS branch
                if ($null -eq $best -or $fixed -lt $best.Fixed) {
                    $desc = Get-AbField $rem 'Description'
                    $best = @{ Fixed = $fixed; KB = [string](Get-AbField $desc 'Value') }
                }
            }
            if (-not $best -or $CurrentBuild -ge $best.Fixed) { continue }
            $severity = ''; $exploited = $false
            foreach ($t in @(Get-AbField $v 'Threats')) {
                $type = [string](Get-AbField $t 'Type'); $text = [string](Get-AbField (Get-AbField $t 'Description') 'Value')
                if ($type -in '3', 'Severity') {
                    $tp = @(Get-AbField $t 'ProductID' | ForEach-Object { [string]$_ })
                    if (-not $tp -or @($tp | Where-Object { $pids -contains $_ })) { $severity = $text }
                }
                elseif ($type -in '1', 'Exploit Status' -and $text -match 'Exploited:\s*Yes') { $exploited = $true }
            }
            $score = 0.0
            foreach ($s in @(Get-AbField $v 'CVSSScoreSets')) { $b = Get-AbField $s 'BaseScore'; if ($b -and [double]$b -gt $score) { $score = [double]$b } }
            $seen[$cve] = $true
            [pscustomobject]@{
                Cve = $cve; Title = [string](Get-AbField (Get-AbField $v 'Title') 'Value'); Severity = $severity; Score = $score
                Exploited = $exploited; Kev = $kev.Contains($cve); FixedBuild = $best.Fixed.ToString(); KB = $best.KB
            }
        }
    }
}

function Test-AbCvrfBuildCovered {
    # True when the feed has at least one fix for this product on this exact build branch
    # (e.g. 10.0.26100.*). Insider builds share a product name with GA releases but never match.
    param([object[]]$Documents, [Parameter(Mandatory)][string[]]$ProductNames, [Parameter(Mandatory)][version]$CurrentBuild)
    foreach ($doc in $Documents) {
        $pids = @(Get-AbCvrfProductId -Node (Get-AbField $doc 'ProductTree') -Names $ProductNames | Sort-Object -Unique)
        if (-not $pids) { continue }
        foreach ($v in @(Get-AbField $doc 'Vulnerability')) {
            foreach ($rem in @(Get-AbField $v 'Remediations')) {
                $fixed = $null
                if (-not [version]::TryParse(([string](Get-AbField $rem 'FixedBuild') -replace '[^\d.]', ''), [ref]$fixed)) { continue }
                if ($fixed.Major -ne $CurrentBuild.Major -or $fixed.Build -ne $CurrentBuild.Build) { continue }
                if (@(Get-AbField $rem 'ProductID' | Where-Object { $pids -contains [string]$_ }).Count) { return $true }
            }
        }
    }
    $false
}

function Get-AbCvrfDocument {
    # Latest $Months monthly MSRC security update documents (cached on disk). Empty when offline.
    param([string]$BaseUrl, [int]$Months = 3)
    $list = Get-AbWebText -Uri "$BaseUrl/updates" -CacheName 'msrc-updates.json' -MaxAgeHours 12
    if (-not $list) { return }
    $items = @(Get-AbField (ConvertFrom-AbJson -Json $list) 'value') |
        Where-Object { (Get-AbField $_ 'ID') -match '^\d{4}-[A-Za-z]{3}$' } |
        Sort-Object { [datetime](Get-AbField $_ 'InitialReleaseDate') } -Descending | Select-Object -First $Months
    foreach ($i in $items) {
        $id = [string](Get-AbField $i 'ID')
        $url = [string](Get-AbField $i 'CvrfUrl')
        if (-not $url) { $url = "$BaseUrl/cvrf/$id" }
        $text = Get-AbWebText -Uri $url -CacheName "msrc-cvrf-$id.json" -MaxAgeHours 24
        if ($text) { , (ConvertFrom-AbJson -Json $text) }
    }
}

function Get-AbKevCve {
    # CVE ids on CISA's Known Exploited Vulnerabilities catalogue (vendor Microsoft unless -AllVendors).
    param([string]$Url, [switch]$AllVendors)
    $text = Get-AbWebText -Uri $Url -CacheName 'cisa-kev.json' -MaxAgeHours 24
    if (-not $text) { return }
    foreach ($v in @(Get-AbField (ConvertFrom-AbJson -Json $text) 'vulnerabilities')) {
        if ($AllVendors -or [string](Get-AbField $v 'vendorProject') -eq 'Microsoft') { [string](Get-AbField $v 'cveID') }
    }
}
