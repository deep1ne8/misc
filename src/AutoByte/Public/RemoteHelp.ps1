function Start-AbRemoteHelp {
    <#
    .SYNOPSIS
        Opens a way for IT support to connect to this PC remotely.
    .DESCRIPTION
        1. If the MSP set RemoteSupportUrl (config.json or the AB_RemoteSupportUrl RMM variable), opens
           that https:// page in the browser (e.g. the MSP's ScreenConnect, Splashtop SOS or support portal).
        2. Otherwise opens Microsoft Quick Assist (built into Windows 10/11).
        It also returns what to tell the technician (computer and user name) and whether the Datto RMM
        agent is installed, in which case the technician can connect with Datto Web Remote directly.
    .PARAMETER Method
        Auto (default), Url or QuickAssist.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([ValidateSet('Auto', 'Url', 'QuickAssist')][string]$Method = 'Auto')
    $cfg = Get-AbConfig
    $url = [string]$cfg.RemoteSupportUrl
    $name = if ($cfg.RemoteSupportName) { [string]$cfg.RemoteSupportName } else { 'your IT support' }
    $datto = [bool](Get-AbService -Name 'CagService')
    if ($Method -eq 'Auto') { $Method = if ($url) { 'Url' } else { 'QuickAssist' } }
    $result = [pscustomobject]@{
        Method = $Method; Target = ''; Opened = $false; SupportName = $name
        Computer = [Environment]::MachineName; User = [Environment]::UserName; DattoAgent = $datto; Message = ''
    }
    if ($Method -eq 'Url') {
        $u = $null
        if (-not $url -or -not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$u) -or $u.Scheme -ne 'https') {
            throw "RemoteSupportUrl must be an https:// address (got '$url'). Set it in %ProgramData%\AutoByte\config.json or the AB_RemoteSupportUrl variable."
        }
        $result.Target = $u.AbsoluteUri
        if ($PSCmdlet.ShouldProcess($u.AbsoluteUri, 'Open the remote support page')) { Start-Process -FilePath $u.AbsoluteUri; $result.Opened = $true }
        $result.Message = "The $name remote support page is opening in your browser."
    }
    else {
        $win = if ($env:SystemRoot) { $env:SystemRoot } else { 'C:\Windows' }
        $legacy = [IO.Path]::Combine($win, 'System32', 'quickassist.exe')
        $result.Target = 'ms-quick-assist:'
        if ($PSCmdlet.ShouldProcess('Quick Assist', 'Open')) {
            try { Start-Process -FilePath 'ms-quick-assist:' -ErrorAction Stop; $result.Opened = $true }
            catch {
                if (Test-AbPath -Path $legacy) { Start-Process -FilePath $legacy; $result.Target = $legacy; $result.Opened = $true }
                else { throw 'Quick Assist is not installed. Get it free from the Microsoft Store ("Quick Assist"), or ask your IT support for their remote support link.' }
            }
        }
        $result.Message = 'Quick Assist is opening. Choose "Get help", then type the 6-digit code your technician gives you.'
    }
    Write-AbLog -Message "Remote help: $($result.Method) $($result.Target) (opened: $($result.Opened))"
    $result
}
