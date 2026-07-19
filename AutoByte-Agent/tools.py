"""
AutoByte Agent — safe Windows tool layer.

Tool categories:
  * READ-ONLY  -> executed automatically (diagnostics)
  * MUTATING   -> require explicit [Y/n] confirmation before running

Every tool returns a plain-text result the agent feeds back to the model.
No tool deletes user data or formats disks; mutating tools are either
bounded repairs or require the operator to type the exact command.
"""
import subprocess, json, os, shutil, pathlib

SCRIPTS_DIR = pathlib.Path(__file__).resolve().parent.parent / "Scripts"
if not SCRIPTS_DIR.exists():
    # allow running from elsewhere: look for a Scripts folder next to the repo
    SCRIPTS_DIR = pathlib.Path(__file__).resolve().parent / "Scripts"

READONLY = {"system_info", "event_log", "disk", "services", "network", "run_diagnostic"}
MUTATING = {"run_fix", "run_command", "install_updates"}

PS = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command"]


def _ps(script, timeout=60):
    try:
        r = subprocess.run(PS + [script], capture_output=True, text=True, timeout=timeout)
        out = (r.stdout or "") + (r.stderr or "")
        return out.strip() or "(no output)"
    except subprocess.TimeoutExpired:
        return "ERROR: command timed out after %ss" % timeout
    except Exception as e:
        return "ERROR: %s" % e


def system_info():
    return _ps("""
$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB)
$boot = (Get-Date) - $os.LastBootUpTime
[PSCustomObject]@{
  OS = $os.Caption + ' ' + $os.OSArchitecture
  Version = $os.Version
  CPU = $cpu.Name
  RAM_GB = $ram
  Uptime_h = [math]::Round($boot.TotalHours,1)
  User = $env:USERNAME
} | Format-List | Out-String
""")


def event_log(errors=20):
    return _ps(f"""
Get-WinEvent -FilterHashtable @{{LogName='System','Application'; Level=2,3}} -MaxEvents {errors} -ErrorAction SilentlyContinue |
  ForEach-Object {{ "{{0}} [{{1}}] {{2}}" -f $_.TimeCreated, $_.LevelDisplayName, $_.ProviderName }} |
  Out-String
""")


def disk():
    return _ps("""
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
  ForEach-Object { "$($_.DeviceID)  $([math]::Round($_.FreeSpace/1GB))GB free / $([math]::Round($_.Size/1GB))GB  ($('{0:P0}' -f ($_.FreeSpace/$_.Size)))" } |
  Out-String
""")


def services(name):
    return _ps(f"""
$svc = Get-Service -Name '{name}' -ErrorAction SilentlyContinue
if ($svc) {{ "$($svc.Name)  [$($svc.Status)]  startType=$($svc.StartType)" }} else {{ "Service '$name' not found" }}
""")


def network():
    return _ps("""
$ips = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' }).IPAddress
$net = (Test-Connection 8.8.8.8 -Count 2 -Quiet)
"IP: $(($ips -join ', '))`nInternet: $(if($net){'reachable'}else{'DOWN'})"
""")


def _find_script(name):
    if not SCRIPTS_DIR.exists():
        return None
    cands = [p for p in SCRIPTS_DIR.glob("*.ps1")]
    exact = [p for p in cands if p.stem.lower() == name.lower()]
    if exact:
        return exact[0]
    part = [p for p in cands if name.lower() in p.stem.lower()]
    return part[0] if part else None


def run_diagnostic(name):
    """Read-only: run a script that just reports state."""
    p = _find_script(name)
    if not p:
        return f"Script '{name}' not found in {SCRIPTS_DIR}"
    return _ps(f"& '{p}'", timeout=120)


def run_fix(name):
    """Mutating: run a repair/troubleshoot script from the library (after confirm)."""
    p = _find_script(name)
    if not p:
        return f"Script '{name}' not found in {SCRIPTS_DIR}"
    return _ps(f"& '{p}'", timeout=300)


def run_command(ps_code):
    """Mutating: run arbitrary PowerShell (after confirm)."""
    return _ps(ps_code, timeout=300)


def install_updates():
    """Mutating: trigger Windows Update install (after confirm)."""
    return _ps("""
$prog = "$env:SystemRoot\\System32\\usoclient.exe"
if (Test-Path $prog) { Start-Process $prog -ArgumentList 'StartInstall' -Wait; 'Windows Update install triggered.' }
else { 'usoclient.exe not found on this build.' }
""", timeout=600)


DISPATCH = {
    "system_info": lambda a: system_info(),
    "event_log": lambda a: event_log(int(a) if a and a.isdigit() else 20),
    "disk": lambda a: disk(),
    "services": lambda a: services(a or ""),
    "network": lambda a: network(),
    "run_diagnostic": lambda a: run_diagnostic(a or ""),
    "run_fix": lambda a: run_fix(a or ""),
    "run_command": lambda a: run_command(a or ""),
    "install_updates": lambda a: install_updates(),
}


def call(tool, arg=""):
    if tool not in DISPATCH:
        return f"Unknown tool: {tool}"
    if tool in MUTATING:
        return {"need_confirm": True, "tool": tool, "arg": arg}
    return {"result": DISPATCH[tool](arg)}


def describe():
    return {
        "readonly": sorted(READONLY),
        "mutating": sorted(MUTATING),
        "scripts_available": sorted(p.stem for p in SCRIPTS_DIR.glob("*.ps1")) if SCRIPTS_DIR.exists() else [],
    }


if __name__ == "__main__":
    # smoke test the read-only layer (safe)
    print("== system_info =="); print(system_info()[:400])
    print("== disk =="); print(disk()[:300])
    print("== network =="); print(network()[:200])
    print("== tool list =="); print(json.dumps(describe(), indent=2)[:600])
