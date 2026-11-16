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
    param([Parameter(Mandatory)][ValidateSet('WindowsTemp', 'UserTemp', 'Spool', 'SoftwareDistribution', 'Catroot2', 'MpCmdRun', 'W32tm', 'IpConfig', 'Sfc')][string]$Name)
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
