# System adapters. Tasks must touch the machine ONLY through these functions, so that every
# task can be unit-tested with Pester mocks (Mock -ModuleName AutoByte <Adapter>) on any OS.
# Keep them thin: no logic beyond calling the underlying command and normalising the result.

function Get-AbCim {
    param([Parameter(Mandatory)][string]$ClassName, [string]$Namespace = 'root/cimv2', [string]$Filter)
    $p = @{ ClassName = $ClassName; Namespace = $Namespace; ErrorAction = 'Stop' }
    if ($Filter) { $p.Filter = $Filter }
    Get-CimInstance @p
}

function Get-AbRegistryValue {
    # Returns the value, or $null if the key/value does not exist.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    $item.$Name
}

function Test-AbPath {
    param([Parameter(Mandatory)][string]$Path)
    Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue
}

function Get-AbService {
    param([string[]]$Name)
    if ($Name) { Get-Service -Name $Name -ErrorAction SilentlyContinue } else { Get-Service -ErrorAction SilentlyContinue }
}

function Set-AbServiceState {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][ValidateSet('Running', 'Stopped')][string]$State)
    if ($State -eq 'Running') { Start-Service -Name $Name -ErrorAction Stop } else { Stop-Service -Name $Name -Force -ErrorAction Stop }
}

function Invoke-AbNative {
    # Runs an executable with an argument list; returns exit code and combined output. Never uses a shell.
    param([Parameter(Mandatory)][string]$FilePath, [string[]]$ArgumentList = @(), [int]$TimeoutSeconds = 1800)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = ($ArgumentList | ForEach-Object { if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ } }) -join ' '
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEndAsync()
    $stderr = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        try { $proc.Kill() } catch { Write-Verbose "Kill failed: $_" }
        return [pscustomobject]@{ ExitCode = 124; Output = "Timed out after $TimeoutSeconds s" }
    }
    [pscustomobject]@{ ExitCode = $proc.ExitCode; Output = ($stdout.Result + $stderr.Result).Trim() }
}

function Get-AbWinEvent {
    param([Parameter(Mandatory)][hashtable]$FilterHashtable, [int]$MaxEvents = 500)
    Get-WinEvent -FilterHashtable $FilterHashtable -MaxEvents $MaxEvents -ErrorAction SilentlyContinue
}

function Get-AbDefenderStatus { Get-MpComputerStatus -ErrorAction Stop }

function Get-AbPhysicalDisk { Get-PhysicalDisk -ErrorAction Stop }

function Get-AbSystemDiskMediaType {
    # 'SSD', 'HDD', 'SCM' or 'Unspecified' for the disk holding the Windows volume; $null if unknown.
    try {
        $letter = ([string]$env:SystemDrive).TrimEnd(':')
        if (-not $letter) { $letter = 'C' }
        $num = (Get-Partition -DriveLetter $letter -ErrorAction Stop | Select-Object -First 1).DiskNumber
        [string](Get-PhysicalDisk -ErrorAction Stop | Where-Object { [string]$_.DeviceId -eq [string]$num } | Select-Object -First 1).MediaType
    }
    catch { Write-Verbose "System disk type unknown: $_"; $null }
}

function Resolve-AbDns {
    param([Parameter(Mandatory)][string]$Name)
    try { [bool]([System.Net.Dns]::GetHostAddresses($Name)) } catch { $false }
}

function Test-AbTcp {
    param([Parameter(Mandatory)][string]$HostName, [int]$Port = 443, [int]$TimeoutMs = 3000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $client.BeginConnect($HostName, $Port, $null, $null)
        $ok = $ar.AsyncWaitHandle.WaitOne($TimeoutMs) -and $client.Connected
        $ok
    }
    catch { $false }
    finally { $client.Close() }
}

function Get-AbChildItem {
    param([Parameter(Mandatory)][string]$Path, [switch]$Recurse)
    # Missing folders yield nothing (the -File switch only binds for existing filesystem paths).
    if (-not (Test-Path -LiteralPath $Path -PathType Container -ErrorAction SilentlyContinue)) { return }
    Get-ChildItem -LiteralPath $Path -Force -File -Recurse:$Recurse -ErrorAction SilentlyContinue
}

function Remove-AbFile {
    # Deletes one file; returns $true if removed. Locked/in-use files are skipped, never forced.
    param([Parameter(Mandatory)][string]$Path)
    try { Remove-Item -LiteralPath $Path -Force -ErrorAction Stop; $true } catch { $false }
}

function Rename-AbItem {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$NewName)
    Rename-Item -LiteralPath $Path -NewName $NewName -ErrorAction Stop
}

function Get-AbDate { Get-Date }

function Test-AbIsAdmin {
    if ($env:OS -ne 'Windows_NT') { return $false }
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-AbServiceStartType {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][ValidateSet('Automatic', 'Manual')][string]$StartupType)
    Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
}

function Get-AbImageHealth {
    # Returns 'Healthy', 'Repairable' or 'NonRepairable' (locale-independent, unlike DISM.exe text).
    "$((Repair-WindowsImage -Online -CheckHealth -ErrorAction Stop).ImageHealthState)"
}

function Repair-AbImage {
    $r = Repair-WindowsImage -Online -RestoreHealth -NoRestart -ErrorAction Stop
    [pscustomobject]@{ State = "$($r.ImageHealthState)"; RestartNeeded = [bool]$r.RestartNeeded }
}

function Get-AbSystemPath {
    # Well-known folders, centralised so tests can redirect them. [IO.Path]::Combine avoids
    # provider drive validation (Join-Path fails if the drive does not exist in the session).
    param([Parameter(Mandatory)][ValidateSet('WindowsTemp', 'UserTemp', 'Spool', 'SoftwareDistribution', 'Catroot2', 'MpCmdRun', 'W32tm', 'IpConfig', 'Sfc', 'PowerCfg', 'PnpUtil')][string]$Name)
    $win = if ($env:SystemRoot) { $env:SystemRoot } else { 'C:\Windows' }
    $pf = if ($env:ProgramFiles) { $env:ProgramFiles } else { 'C:\Program Files' }
    switch ($Name) {
        'WindowsTemp' { [IO.Path]::Combine($win, 'Temp') }
        'UserTemp' { [IO.Path]::GetTempPath().TrimEnd('\', '/') }
        'Spool' { [IO.Path]::Combine($win, 'System32', 'spool', 'PRINTERS') }
        'SoftwareDistribution' { [IO.Path]::Combine($win, 'SoftwareDistribution') }
        'Catroot2' { [IO.Path]::Combine($win, 'System32', 'catroot2') }
        'MpCmdRun' { [IO.Path]::Combine($pf, 'Windows Defender', 'MpCmdRun.exe') }
        'W32tm' { [IO.Path]::Combine($win, 'System32', 'w32tm.exe') }
        'IpConfig' { [IO.Path]::Combine($win, 'System32', 'ipconfig.exe') }
        'Sfc' { [IO.Path]::Combine($win, 'System32', 'sfc.exe') }
        'PowerCfg' { [IO.Path]::Combine($win, 'System32', 'powercfg.exe') }
        'PnpUtil' { [IO.Path]::Combine($win, 'System32', 'pnputil.exe') }
    }
}

function Get-AbTimeOffset {
    # SNTP query; returns (server time - local time) in seconds, or $null if unreachable.
    # Locale-independent (no parsing of w32tm text output).
    param([string]$Server = 'time.windows.com', [int]$TimeoutMs = 3000)
    $udp = New-Object System.Net.Sockets.UdpClient
    try {
        $udp.Client.ReceiveTimeout = $TimeoutMs
        $req = New-Object byte[] 48
        $req[0] = 0x1B
        $t1 = [DateTime]::UtcNow
        [void]$udp.Send($req, $req.Length, $Server, 123)
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $resp = $udp.Receive([ref]$ep)
        $t4 = [DateTime]::UtcNow
        if ($resp.Length -lt 48) { return $null }
        [uint64]$intPart = ([uint64]$resp[40] -shl 24) -bor ([uint64]$resp[41] -shl 16) -bor ([uint64]$resp[42] -shl 8) -bor [uint64]$resp[43]
        [uint64]$fracPart = ([uint64]$resp[44] -shl 24) -bor ([uint64]$resp[45] -shl 16) -bor ([uint64]$resp[46] -shl 8) -bor [uint64]$resp[47]
        $ms = ($intPart * 1000) + (($fracPart * 1000) / 0x100000000)
        $server = (New-Object DateTime(1900, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)).AddMilliseconds([double]$ms)
        $local = $t1.AddTicks([long](($t4 - $t1).Ticks / 2))
        [math]::Round(($server - $local).TotalSeconds, 2)
    }
    catch { $null }
    finally { $udp.Close() }
}

# ---- event logs -----------------------------------------------------------------------------
function Get-AbEventRecord {
    <#
      Returns normalised events: Log, Provider, Id, Level, Time, Data (named EventData fields,
      locale-independent) and Props (positional insertion strings). Missing/disabled/unauthorised
      logs return nothing.
    #>
    param([Parameter(Mandatory)][string]$LogName, [int[]]$Id, [Parameter(Mandatory)][datetime]$StartTime, [int]$MaxEvents = 2000)
    $filter = @{ LogName = $LogName; StartTime = $StartTime }
    if ($Id) { $filter.Id = $Id }
    try { $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEvents -ErrorAction Stop) } catch { return }
    foreach ($e in $events) {
        $data = @{}
        try {
            $xml = [xml]$e.ToXml()
            foreach ($d in @($xml.SelectNodes("//*[local-name()='EventData']/*[local-name()='Data']"))) {
                $n = $d.GetAttribute('Name')
                if ($n) { $data[$n] = $d.InnerText }
            }
            foreach ($d in @($xml.SelectNodes("//*[local-name()='UserData']//*[not(*)]"))) { $data[$d.LocalName] = $d.InnerText }
        }
        catch { Write-Verbose "Event XML unreadable: $_" }
        [pscustomobject]@{
            Log = $LogName; Provider = [string]$e.ProviderName; Id = [int]$e.Id; Level = [int]$e.Level
            Time = $e.TimeCreated; Data = $data; Props = @($e.Properties | ForEach-Object { [string]$_.Value })
        }
    }
}

# ---- Windows Update Agent (COM) ------------------------------------------------------------
function Get-AbPendingUpdate {
    # Updates not yet installed, as reported by the configured source (Windows Update, WSUS, Intune).
    param([string]$Criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'")
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'AutoByte'
    $result = $session.CreateUpdateSearcher().Search($Criteria)
    foreach ($u in $result.Updates) {
        [pscustomobject]@{
            UpdateId    = [string]$u.Identity.UpdateID
            Title       = [string]$u.Title
            KB          = (@($u.KBArticleIDs) | ForEach-Object { "KB$_" }) -join ','
            Severity    = [string]$u.MsrcSeverity
            CategoryIds = @($u.Categories | ForEach-Object { ([string]$_.CategoryID).ToLowerInvariant() })
            Type        = [int]$u.Type
            SizeMB      = [math]::Round([double]$u.MaxDownloadSize / 1MB, 1)
            Reboot      = [bool]$u.InstallationBehavior.RebootBehavior
        }
    }
}

function Install-AbUpdate {
    # Downloads and installs the given update ids. Returns per-update result codes and reboot flag.
    param([Parameter(Mandatory)][string[]]$UpdateId, [string]$Criteria = "IsInstalled=0 and IsHidden=0")
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'AutoByte'
    $found = $session.CreateUpdateSearcher().Search($Criteria).Updates
    $coll = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $found) {
        if ($UpdateId -contains [string]$u.Identity.UpdateID) {
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            [void]$coll.Add($u)
        }
    }
    if ($coll.Count -eq 0) { return [pscustomobject]@{ Installed = 0; Failed = 0; RebootRequired = $false; Results = @() } }
    $dl = $session.CreateUpdateDownloader(); $dl.Updates = $coll; [void]$dl.Download()
    $inst = $session.CreateUpdateInstaller(); $inst.Updates = $coll
    $res = $inst.Install()
    $rows = for ($i = 0; $i -lt $coll.Count; $i++) {
        $r = $res.GetUpdateResult($i)
        [pscustomobject]@{ Title = [string]$coll.Item($i).Title; ResultCode = [int]$r.ResultCode; HResult = ('0x{0:X8}' -f $r.HResult) }
    }
    [pscustomobject]@{
        Installed      = @($rows | Where-Object { $_.ResultCode -in 2, 3 }).Count
        Failed         = @($rows | Where-Object { $_.ResultCode -in 4, 5 }).Count
        RebootRequired = [bool]$res.RebootRequired
        Results        = @($rows)
    }
}

# ---- web data with on-disk cache (CVE feeds) -------------------------------------------------
function Get-AbCachePath {
    param([Parameter(Mandatory)][string]$Name)
    $base = if ($env:ProgramData) { $env:ProgramData } else { [IO.Path]::GetTempPath() }
    $dir = [IO.Path]::Combine($base, 'AutoByte', 'Cache')
    try { if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop } }
    catch { $dir = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'AutoByte-Cache'); $null = New-Item -ItemType Directory -Path $dir -Force }  # standard user: ProgramData cache owned by SYSTEM
    [IO.Path]::Combine($dir, $Name)
}

function Get-AbWebText {
    # GET a URL as text, cached on disk for MaxAgeHours. Returns $null when offline and no cache.
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$CacheName, [int]$MaxAgeHours = 24, [string]$Accept = 'application/json')
    $cache = Get-AbCachePath -Name $CacheName
    if ((Test-Path -LiteralPath $cache) -and ((Get-Date) - (Get-Item -LiteralPath $cache).LastWriteTime).TotalHours -lt $MaxAgeHours) {
        return [IO.File]::ReadAllText($cache)
    }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $r = Invoke-WebRequest -Uri $Uri -UseBasicParsing -Headers @{ Accept = $Accept } -TimeoutSec 120 -ErrorAction Stop
        $text = if ($r.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($r.Content) } else { [string]$r.Content }
        try { [IO.File]::WriteAllText($cache, $text) } catch { Write-Verbose "Cache not writable: $_" }
        $text
    }
    catch {
        if (Test-Path -LiteralPath $cache) { return [IO.File]::ReadAllText($cache) }  # stale cache beats nothing
        $null
    }
}

function ConvertFrom-AbJson {
    # Large-document JSON parser. Windows PowerShell 5.1's ConvertFrom-Json is limited to ~2 MB,
    # so Desktop uses JavaScriptSerializer (dictionaries/arrays). Read fields with Get-AbField.
    param([Parameter(Mandatory)][string]$Json)
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        Add-Type -AssemblyName System.Web.Extensions
        $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $js.MaxJsonLength = [int]::MaxValue
        $js.RecursionLimit = 256
        return , $js.DeserializeObject($Json)
    }
    ConvertFrom-Json -InputObject $Json
}

function Get-AbField {
    # Reads a field from a PSObject or dictionary; $null when absent (StrictMode-safe).
    param($Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if (([System.Collections.IDictionary]$Object).Contains($Name)) { return $Object[$Name] } else { return $null } }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { $p.Value } else { $null }
}

# ---- winget --------------------------------------------------------------------------------
function Get-AbWingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    # Running as SYSTEM (RMM): winget is not on PATH; use the newest App Installer package.
    $pf = if ($env:ProgramFiles) { $env:ProgramFiles } else { 'C:\Program Files' }
    $root = [IO.Path]::Combine($pf, 'WindowsApps')
    if (-not (Test-Path -LiteralPath $root)) { return $null }
    $dirs = @(Get-ChildItem -LiteralPath $root -Directory -Filter 'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    foreach ($d in $dirs) { $exe = [IO.Path]::Combine($d.FullName, 'winget.exe'); if (Test-Path -LiteralPath $exe) { return $exe } }
    $null
}

function Get-AbAppUpgrade {
    # Apps with an available upgrade: Name, Id, Version, Available (an array, possibly empty).
    # $null only when winget itself is unavailable.
    $winget = Get-AbWingetPath
    if (-not $winget) { return $null }
    $r = Invoke-AbNative -FilePath $winget -ArgumentList 'upgrade', '--include-unknown', '--accept-source-agreements', '--disable-interactivity' -TimeoutSeconds 300
    , @(ConvertFrom-AbWingetTable -Text $r.Output)
}

function ConvertFrom-AbWingetTable {
    # Parses winget's table by column positions taken from the header row (works in any UI language).
    param([string]$Text)
    $lines = @($Text -split "`r?`n" | ForEach-Object { ($_ -replace '^.*\r', '').TrimEnd() })
    $sep = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^-{10,}$') { $sep = $i; break } }
    if ($sep -lt 1 -or $sep + 1 -ge $lines.Count) { return @() }
    $header = $lines[$sep - 1]
    $starts = @([regex]::Matches($header, '(?<=^|\s)\S') | ForEach-Object Index)
    if ($starts.Count -lt 4) { return @() }
    $starts = $starts[0..([Math]::Min(4, $starts.Count - 1))]
    foreach ($line in $lines[($sep + 1)..($lines.Count - 1)]) {
        if ($line.Length -le $starts[3] -or $line -match '^\d+ upgrades? available' -or $line -notmatch '\S') { continue }
        $cell = { param($n) $s = $starts[$n]; $e = if ($n + 1 -lt $starts.Count) { $starts[$n + 1] } else { $line.Length }; if ($s -ge $line.Length) { '' } else { $line.Substring($s, [Math]::Min($e, $line.Length) - $s).Trim() } }
        $id = & $cell 1
        if ($id -notmatch '^[A-Za-z0-9][A-Za-z0-9._+\-]*$') { continue }
        [pscustomobject]@{ Name = (& $cell 0); Id = $id; Version = (& $cell 2); Available = (& $cell 3) }
    }
}

# ---- misc ----------------------------------------------------------------------------------
function Wait-AbInterval { param([int]$Milliseconds) Start-Sleep -Milliseconds $Milliseconds }

function Get-AbOsBuild {
    # Exact OS build incl. revision (UBR), e.g. 10.0.26100.2894, plus naming facts for CVE matching.
    $k = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    [pscustomobject]@{
        Build            = [string](Get-AbRegistryValue -Path $k -Name 'CurrentBuild')
        Ubr              = [int](Get-AbRegistryValue -Path $k -Name 'UBR')
        DisplayVersion   = [string](Get-AbRegistryValue -Path $k -Name 'DisplayVersion')
        ProductName      = [string](Get-AbRegistryValue -Path $k -Name 'ProductName')
        InstallationType = [string](Get-AbRegistryValue -Path $k -Name 'InstallationType')
        Architecture     = [string]$env:PROCESSOR_ARCHITECTURE
    }
}

function Get-AbVolumeScan {
    # Online file-system scan: NoErrorsFound | SpotFixNeeded | FullRepairNeeded | ... (locale-independent).
    param([string]$DriveLetter = 'C')
    "$(Repair-Volume -DriveLetter $DriveLetter -Scan -ErrorAction Stop)"
}

function Repair-AbVolume {
    param([string]$DriveLetter = 'C', [ValidateSet('SpotFix', 'OfflineScanAndFix')][string]$Mode = 'SpotFix')
    if ($Mode -eq 'SpotFix') { "$(Repair-Volume -DriveLetter $DriveLetter -SpotFix -ErrorAction Stop)" }
    else { "$(Repair-Volume -DriveLetter $DriveLetter -OfflineScanAndFix -ErrorAction Stop)" }
}
