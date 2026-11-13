"""AutoByte Agent - safe Windows tool layer.

Security model (tool arguments come from the LLM and are treated as hostile):

* READ-ONLY tools run automatically. ``run_diagnostic`` may only run scripts the
  generated manifest marks ``agentDiagnostic`` (non-mutating, non-interactive,
  supported), matched by exact name and hash-verified before execution.
* MUTATING tools return ``need_confirm``; the operator must approve, then
  :func:`execute` re-validates everything before running.
* ``run_command`` is refused outright when it matches the destructive deny-list.
* Every proposal, decision and mutating execution is written to the audit log.
* Output is capped; every subprocess has a timeout; no shell=True anywhere.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
import pathlib
import re
import shutil
import subprocess
import time
from collections.abc import Callable
from typing import Any, Final, TypedDict, cast

HERE: Final = pathlib.Path(__file__).resolve().parent
REPO_ROOT: Final = HERE.parent
SCRIPTS_DIR: Final = REPO_ROOT / "Scripts"
MANIFEST_PATH: Final = REPO_ROOT / "scripts.manifest.json"

READONLY: Final = frozenset({"system_info", "event_log", "disk", "services", "network", "run_diagnostic"})
MUTATING: Final = frozenset({"run_fix", "run_command", "install_updates"})

MAX_OUTPUT_CHARS: Final = 20_000
MAX_COMMAND_CHARS: Final = 4_000
SERVICE_NAME_RE: Final = re.compile(r"^[A-Za-z0-9_.\-$]{1,256}$")
SCRIPT_NAME_RE: Final = re.compile(r"^[A-Za-z0-9_.\-]{1,128}$")

# Destructive or evasive PowerShell patterns. Matching commands are refused, not
# offered for approval: one mistyped "y" must not be able to wipe a client PC.
DENY_PATTERNS: Final[tuple[tuple[str, re.Pattern[str]], ...]] = tuple(
    (label, re.compile(rx, re.IGNORECASE))
    for label, rx in (
        (
            "disk wipe/format",
            r"\b(Format-Volume|Clear-Disk|Initialize-Disk|Remove-Partition|Set-Disk|diskpart|format(\.com)?\s+[a-z]:)",
        ),
        ("boot config", r"\b(bcdedit|bcdboot|bootsect|reagentc)\b"),
        ("secure wipe", r"\bcipher(\.exe)?\s+/w\b|\bsdelete\b"),
        (
            "registry hive delete",
            r"\b(Remove-Item|ri)\b[^\n;|]*(HKLM|HKCU|HKU|Registry::)[:\\]*\\?(SOFTWARE|SYSTEM|SAM|SECURITY)?\\?['\"]?\s*(-Recurse|$)|\breg(\.exe)?\s+delete\s+(HKLM|HKEY_LOCAL_MACHINE)\\(SOFTWARE|SYSTEM)\\?\s",
        ),
        (
            "shadow copy / backup destruction",
            r"\b(vssadmin|wbadmin)\b[^\n]*\bdelete\b|Win32_ShadowCopy[^\n]*Delete",
        ),
        (
            "disable security tooling",
            r"Set-MpPreference[^\n]*-Disable\w+\s+\$?true|Uninstall-WindowsFeature[^\n]*Defender",
        ),
        (
            "account/credential tampering",
            r"\b(net\s+user\s+\S+\s+\S+|Set-LocalUser[^\n]*-Password|Remove-LocalUser|Disable-LocalUser|Add-LocalGroupMember[^\n]*Administrators)",
        ),
        (
            "download and execute",
            r"\b(Invoke-Expression|iex)\b|DownloadString|DownloadFile[^\n]*;\s*&|Invoke-WebRequest[^\n]*\|\s*(iex|Invoke-Expression)|FromBase64String[^\n]*\b(iex|Invoke-Expression|&)",
        ),
        ("encoded command", r"-(e|en|enc|enco|encod|encode|encoded|encodedcommand)\s+[A-Za-z0-9+/=]{16,}"),
        ("execution policy change", r"\bSet-ExecutionPolicy\b"),
        ("forced reboot/shutdown", r"\b(Stop-Computer|Restart-Computer)\b|\bshutdown(\.exe)?\s+/[rsp]"),
    )
)


# Recursive deletes are checked per statement because flag order and aliases vary
# (rm -r -fo, Remove-Item -Recurse X, rd /s /q X). Deep paths such as
# $env:TEMP\\x or C:\\Users\\bob\\Downloads stay allowed (with approval);
# the protected roots themselves are refused.
_DELETE_VERB: Final = re.compile(r"(?:^|[\s(&{])(Remove-Item|rm|rmdir|rd|del|erase|ri)(?=\s)", re.IGNORECASE)
_DOTNET_DELETE: Final = re.compile(r"(\.|::)(Delete|DeleteSubKeyTree)\s*\(", re.IGNORECASE)
_RECURSE_FLAG: Final = re.compile(r"(?:\s|^)(-r\w*|/s)(?=[\s:]|$)", re.IGNORECASE)
_END: Final = r"(?=\\?\*?['\"]?(?:[\s,)}]|$))"
_PROTECTED_ROOT: Final = re.compile(
    r"['\"]?(?:[a-z]:\\?"
    + _END
    + r"|[a-z]:\\(?:windows|program\s?files(?:\s?\(x86\))?|programdata|users|boot|recovery)"
    + _END
    + r"|[a-z]:\\users\\[^\\\s'\"]+"
    + _END
    + r"|\$env:(?:SystemRoot|windir|SystemDrive|ProgramData|ProgramFiles|ProgramW6432|USERPROFILE|HOMEDRIVE|HOMEPATH|PUBLIC|ALLUSERSPROFILE)"
    + _END
    + r"|(?:HKLM|HKCU|HKU|HKCR|HKCC|Registry::HKEY_\w+)(?::|\\)\\?(?:SOFTWARE|SYSTEM|SAM|SECURITY)?"
    + _END
    + r")",
    re.IGNORECASE,
)


def _is_protected_recursive_delete(ps_code: str) -> bool:
    # Split on statements, not pipes: "gci C:\\Users -r | Remove-Item" is one statement.
    for stmt in re.split(r"[;\n]|&&|\|\|", ps_code):
        if not _PROTECTED_ROOT.search(stmt):
            continue
        if _DOTNET_DELETE.search(stmt):
            return True
        if _DELETE_VERB.search(stmt) and _RECURSE_FLAG.search(stmt):
            return True
    return False


class ToolResult(TypedDict, total=False):
    status: str  # ok | need_confirm | refused | error
    tool: str
    arg: str
    result: str
    preview: str


# ---- audit log ---------------------------------------------------------------
def _log_dir() -> pathlib.Path:
    override = os.environ.get("AUTOBYTE_LOG_DIR")
    if override:
        return pathlib.Path(override)
    if os.name == "nt":
        base = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
        return pathlib.Path(base) / "AutoByte" / "Logs"
    return pathlib.Path.home() / ".autobyte" / "logs"


def audit(event: str, **fields: Any) -> None:
    """Append one JSON line to the daily audit log. Never raises."""
    try:
        d = _log_dir()
        d.mkdir(parents=True, exist_ok=True)
        now = _dt.datetime.now(_dt.UTC)
        rec = {"ts": now.isoformat(timespec="seconds"), "event": event, **fields}
        path = d / f"Agent_{now:%Y%m%d}.log"
        with path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass


# ---- process helpers ---------------------------------------------------------
def _powershell() -> list[str]:
    exe = shutil.which("powershell") or shutil.which("pwsh") or "powershell"
    return [exe, "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass"]


def _cap(text: str) -> str:
    text = text.strip() or "(no output)"
    if len(text) > MAX_OUTPUT_CHARS:
        cut = len(text) - MAX_OUTPUT_CHARS
        text = text[:MAX_OUTPUT_CHARS] + f"\n...[truncated {cut} chars]"
    return text


def _run(args: list[str], timeout: int) -> tuple[int, str]:
    try:
        r = subprocess.run(  # noqa: S603 - argument list, no shell
            args,
            capture_output=True,
            text=True,
            timeout=timeout,
            stdin=subprocess.DEVNULL,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except subprocess.TimeoutExpired:
        return 124, f"ERROR: command timed out after {timeout}s"
    except OSError as e:
        return 127, f"ERROR: {e}"
    return r.returncode, _cap((r.stdout or "") + (r.stderr or ""))


def _ps(script: str, timeout: int = 60) -> str:
    code, out = _run([*_powershell(), "-Command", script], timeout)
    return out if code in (0, 124, 127) else f"{out}\n(exit code {code})"


# ---- manifest ----------------------------------------------------------------
class ScriptEntry(TypedDict, total=False):
    name: str
    path: str
    sha256: str
    description: str
    platform: str
    mutating: bool
    requiresAdmin: bool
    interactive: bool
    unsupported: bool
    agentDiagnostic: bool


def load_manifest(path: pathlib.Path | None = None) -> dict[str, ScriptEntry]:
    """Return manifest entries keyed by lower-case name ({} if missing/invalid)."""
    try:
        data = json.loads((path or MANIFEST_PATH).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return {
        str(s["name"]).lower(): cast(ScriptEntry, s)
        for s in data.get("scripts", [])
        if isinstance(s, dict) and "name" in s and "path" in s and "sha256" in s
    }


def normalized_sha256(path: pathlib.Path) -> str:
    """sha256 over CRLF->LF-normalised bytes (matches Build-ScriptManifest.ps1)."""
    return hashlib.sha256(path.read_bytes().replace(b"\r\n", b"\n")).hexdigest()


def resolve_script(
    name: str, *, purpose: str, manifest: dict[str, ScriptEntry] | None = None
) -> tuple[pathlib.Path | None, str]:
    """Validate an LLM-supplied script name. Returns (path, "") or (None, reason)."""
    name = (name or "").strip()
    if name.lower().endswith(".ps1"):
        name = name[:-4]
    if not SCRIPT_NAME_RE.fullmatch(name):
        return None, f"Invalid script name {name!r}: use the exact name from the manifest."
    entries = load_manifest() if manifest is None else manifest
    if not entries:
        return None, "Script manifest missing or unreadable; script tools are disabled."
    entry = entries.get(name.lower())
    if entry is None:
        return None, f"Unknown script {name!r}. Exact names only; no partial matches."
    if entry.get("platform", "windows") != "windows" or not str(entry["path"]).endswith(".ps1"):
        return None, f"{entry['name']} is not a Windows PowerShell script."
    if entry.get("unsupported"):
        return None, f"{entry['name']} is marked unsupported and cannot be run by the agent."
    if purpose == "diagnostic" and not entry.get("agentDiagnostic"):
        return None, (
            f"{entry['name']} is not a read-only diagnostic. "
            "Propose run_fix instead so the operator can approve it."
        )
    if entry.get("interactive"):
        return None, f"{entry['name']} prompts for input; run it from AutoByte-CLI.ps1 instead."
    path = (REPO_ROOT / str(entry["path"])).resolve()
    try:
        path.relative_to(SCRIPTS_DIR.resolve())
    except ValueError:
        return None, "Manifest path escapes the Scripts directory; refusing."
    if not path.is_file():
        return None, f"Script file missing: {entry['path']}"
    if normalized_sha256(path) != str(entry["sha256"]).lower():
        audit("hash_mismatch", script=entry["name"], path=str(path))
        return None, f"Hash mismatch for {entry['name']}; file differs from manifest. Refusing."
    return path, ""


def _run_script(path: pathlib.Path, timeout: int) -> str:
    code, out = _run([*_powershell(), "-File", str(path)], timeout)
    return f"{out}\n(exit code {code})"


# ---- deny-list ---------------------------------------------------------------
def check_command(ps_code: str) -> str | None:
    """Return a refusal reason if ``ps_code`` is not allowed to be proposed."""
    if not ps_code.strip():
        return "Empty command."
    if len(ps_code) > MAX_COMMAND_CHARS:
        return f"Command longer than {MAX_COMMAND_CHARS} characters."
    if "\x00" in ps_code:
        return "Command contains NUL bytes."
    if _is_protected_recursive_delete(ps_code):
        return "Blocked by safety deny-list (recursive delete of a system, user or registry root)."
    for label, rx in DENY_PATTERNS:
        if rx.search(ps_code):
            return f"Blocked by safety deny-list ({label})."
    return None


# ---- read-only tools -----------------------------------------------------------
def system_info() -> str:
    return _ps(
        """
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
"""
    )


def event_log(count: int = 20) -> str:
    count = max(1, min(int(count), 200))
    return _ps(
        f"""
Get-WinEvent -FilterHashtable @{{LogName='System','Application'; Level=2,3}} -MaxEvents {count} -ErrorAction SilentlyContinue |
  ForEach-Object {{ "{{0}} [{{1}}] {{2}} ({{3}}): {{4}}" -f $_.TimeCreated, $_.LevelDisplayName, $_.ProviderName, $_.Id, (($_.Message -split "`n")[0]) }} |
  Out-String -Width 300
"""
    )


def disk() -> str:
    return _ps(
        """
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
  ForEach-Object { "$($_.DeviceID)  $([math]::Round($_.FreeSpace/1GB))GB free / $([math]::Round($_.Size/1GB))GB  ($('{0:P0}' -f ($_.FreeSpace/$_.Size)))" } |
  Out-String
"""
    )


def services(name: str) -> str:
    name = (name or "").strip()
    if not SERVICE_NAME_RE.fullmatch(name):
        return "Invalid service name: letters, digits, '_', '-', '.', '$' only."
    # Name is validated above and contains no quotes, so it is safe inside '...'.
    return _ps(
        f"""
$svc = Get-Service -Name '{name}' -ErrorAction SilentlyContinue
if ($svc) {{ "$($svc.Name)  [$($svc.Status)]  startType=$($svc.StartType)" }} else {{ "Service '{name}' not found" }}
"""
    )


def network() -> str:
    return _ps(
        """
$ips = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' }).IPAddress
$net = (Test-Connection 8.8.8.8 -Count 2 -Quiet)
$dns = try { [bool](Resolve-DnsName microsoft.com -ErrorAction Stop) } catch { $false }
"IP: $(($ips -join ', '))`nInternet: $(if($net){'reachable'}else{'DOWN'})`nDNS: $(if($dns){'ok'}else{'FAILING'})"
"""
    )


def run_diagnostic(name: str) -> str:
    path, why = resolve_script(name, purpose="diagnostic")
    if path is None:
        return why
    audit("diagnostic", script=path.stem)
    return _run_script(path, timeout=120)


# ---- mutating tools (only reachable via execute()) -----------------------------
def run_fix(name: str) -> str:
    path, why = resolve_script(name, purpose="fix")
    if path is None:
        return why
    return _run_script(path, timeout=900)


def run_command(ps_code: str) -> str:
    why = check_command(ps_code)
    if why:
        return f"REFUSED: {why}"
    return _ps(ps_code, timeout=300)


def install_updates() -> str:
    return _ps(
        """
$prog = "$env:SystemRoot\\System32\\usoclient.exe"
if (Test-Path $prog) { Start-Process $prog -ArgumentList 'StartInstall' -Wait; 'Windows Update install triggered.' }
else { 'usoclient.exe not found on this build.' }
""",
        timeout=600,
    )


def _event_count(arg: str) -> int:
    return int(arg) if arg.isdigit() else 20


DISPATCH: Final[dict[str, Callable[[str], str]]] = {
    "system_info": lambda a: system_info(),
    "event_log": lambda a: event_log(_event_count(a)),
    "disk": lambda a: disk(),
    "services": services,
    "network": lambda a: network(),
    "run_diagnostic": run_diagnostic,
    "run_fix": run_fix,
    "run_command": run_command,
    "install_updates": lambda a: install_updates(),
}


def preview(tool: str, arg: str) -> str:
    """Human-readable description of exactly what a mutating call will do."""
    if tool == "run_fix":
        path, _ = resolve_script(arg, purpose="fix")
        return f"Run script {path.name if path else arg} (hash-verified)"
    if tool == "run_command":
        return f"Run PowerShell:\n    {arg}"
    if tool == "install_updates":
        return "Trigger Windows Update install (usoclient StartInstall)"
    return f"{tool} {arg}".strip()


def call(tool: str, arg: str = "") -> ToolResult:
    """First stage: run read-only tools, or validate and hold mutating ones."""
    tool = str(tool or "")
    arg = str(arg or "")
    if tool not in DISPATCH:
        return {"status": "error", "tool": tool, "result": f"Unknown tool: {tool!r}"}
    if tool in MUTATING:
        if tool == "run_command":
            why = check_command(arg)
            if why:
                audit("refused", tool=tool, arg=arg, reason=why)
                return {"status": "refused", "tool": tool, "arg": arg, "result": why}
        if tool == "run_fix":
            path, why = resolve_script(arg, purpose="fix")
            if path is None:
                audit("refused", tool=tool, arg=arg, reason=why)
                return {"status": "refused", "tool": tool, "arg": arg, "result": why}
        audit("proposed", tool=tool, arg=arg)
        return {"status": "need_confirm", "tool": tool, "arg": arg, "preview": preview(tool, arg)}
    return {"status": "ok", "tool": tool, "result": DISPATCH[tool](arg)}


def execute(tool: str, arg: str, *, approved: bool) -> ToolResult:
    """Second stage: run an operator-approved mutating tool, re-validating first."""
    if tool not in MUTATING:
        return {"status": "error", "tool": tool, "result": "execute() is only for mutating tools."}
    if not approved:
        audit("declined", tool=tool, arg=arg)
        return {"status": "refused", "tool": tool, "arg": arg, "result": "Operator declined."}
    held = call(tool, arg)
    if held.get("status") != "need_confirm":
        return held
    start = time.monotonic()
    out = DISPATCH[tool](arg)
    audit("executed", tool=tool, arg=arg, seconds=round(time.monotonic() - start, 1), output_tail=out[-500:])
    return {"status": "ok", "tool": tool, "arg": arg, "result": out}


def describe() -> dict[str, list[str]]:
    m = load_manifest()
    return {
        "readonly": sorted(READONLY),
        "mutating": sorted(MUTATING),
        "diagnostic_scripts": sorted(e["name"] for e in m.values() if e.get("agentDiagnostic")),
        "fix_scripts": sorted(
            e["name"]
            for e in m.values()
            if not e.get("agentDiagnostic")
            and not e.get("unsupported")
            and not e.get("interactive")
            and e.get("platform") == "windows"
        ),
    }


if __name__ == "__main__":
    print(json.dumps(describe(), indent=2))
