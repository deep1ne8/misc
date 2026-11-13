"""AutoByte GUI - Tk front-end that runs pinned, hash-verified scripts.

Scripts are fetched from a release tag (or commit SHA) of the repository, checked
against scripts.manifest.json at the same ref, then executed from a temp file.
Override the ref with the AUTOBYTE_REF environment variable.
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from typing import Any
from urllib.parse import quote

import requests

REPOSITORY = "deep1ne8/misc"
REF = os.environ.get("AUTOBYTE_REF", "v2.0.0")
RAW_BASE = f"https://raw.githubusercontent.com/{REPOSITORY}/{REF}"
PINNED_RE = re.compile(r"^(v\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?|[0-9a-f]{40})$")

# (button label, manifest script name)
GUI_SCRIPTS = [
    ("Disk Cleaner", "DiskCleaner"),
    ("Enable Files On Demand", "EnableFilesOnDemand"),
    ("Download & Install Package", "DownloadandInstallPackage"),
    ("Check User Profile", "CheckUserProfileIssue"),
    ("Dell Bloatware Remover", "BloatWareRemover"),
    ("Reset & Install Windows Update", "InstallWindowsUpdate"),
    ("Windows System Repair", "WindowsSystemRepair"),
    ("Reset Windows Search DB", "ResetandClearWindowsSearchDB"),
    ("Install MS Projects", "InstallMSProjects"),
    ("Check Drive Space", "CheckDriveSpace"),
    ("Internet Speed Test", "InternetSpeedTest"),
    ("Internet Latency Test", "InternetLatencyTest"),
    ("Monitor Troubleshooter", "WorkPaperMonitorTroubleShooter"),
]

# Prepended AFTER hash verification: routes Read-Host to the GUI input box.
PS_OVERRIDE = r"""
function Read-Host {
    param([string]$prompt="")
    Write-Host $prompt -NoNewline
    return [Console]::In.ReadLine()
}
"""


def normalized_sha256(data: bytes) -> str:
    """sha256 over CRLF->LF-normalised bytes (matches build/Build-ScriptManifest.ps1)."""
    return hashlib.sha256(data.replace(b"\r\n", b"\n")).hexdigest()


def fetch_manifest() -> dict[str, dict[str, Any]]:
    resp = requests.get(f"{RAW_BASE}/scripts.manifest.json", timeout=30)
    resp.raise_for_status()
    return {s["name"].lower(): s for s in resp.json()["scripts"]}


def fetch_verified(manifest: dict[str, dict[str, Any]], name: str) -> bytes:
    entry = manifest.get(name.lower())
    if entry is None:
        raise ValueError(f"{name} is not in the manifest for {REF}")
    if entry.get("unsupported"):
        raise ValueError(f"{name} is flagged unsupported")
    resp = requests.get(f"{RAW_BASE}/{quote(entry['path'])}", timeout=60)
    resp.raise_for_status()
    got = normalized_sha256(resp.content)
    if got != entry["sha256"]:
        raise ValueError(f"SHA-256 mismatch for {name}: expected {entry['sha256']}, got {got}")
    return resp.content


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.proc: subprocess.Popen[str] | None = None
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.manifest: dict[str, dict[str, Any]] | None = None

        root.title(f"AutoBytePro GUI  ({REPOSITORY}@{REF})")
        root.geometry("960x680")
        frame = tk.Frame(root)
        frame.pack(fill="x", padx=10, pady=5)
        self.buttons: list[tk.Button] = []
        cols = 7
        for idx, (label, name) in enumerate(GUI_SCRIPTS):
            btn = tk.Button(
                frame,
                text=label,
                width=16,
                height=2,
                wraplength=120,
                command=lambda n=name, d=label: self.run_script(n, d),
            )
            row, col = divmod(idx, cols)
            btn.grid(row=row, column=col, padx=4, pady=4)
            self.buttons.append(btn)
        self.stop_btn = tk.Button(
            frame, text="Stop", width=12, height=2, fg="red", command=self.stop, state="disabled"
        )
        self.stop_btn.grid(row=0, column=cols, padx=(20, 4), pady=4)
        tk.Button(frame, text="Clear", width=12, height=2, command=self.clear).grid(
            row=1, column=cols, padx=(20, 4), pady=4
        )

        self.status = tk.Label(root, text="Ready", anchor="w")
        self.status.pack(fill="x", padx=10, pady=(0, 5))
        self.out = ScrolledText(root, state="disabled", font=("Consolas", 10), bg="black", fg="white")
        self.out.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        inp = tk.Frame(root)
        inp.pack(fill="x", padx=10, pady=(0, 10))
        tk.Label(inp, text="Input:").pack(side="left")
        self.entry = tk.Entry(inp)
        self.entry.pack(side="left", fill="x", expand=True, padx=5)
        self.entry.bind("<Return>", self.send_input)
        tk.Button(inp, text="Send", command=self.send_input).pack(side="left")

        if not PINNED_RE.match(REF):
            self.log(f"[warning] AUTOBYTE_REF={REF!r} is not a release tag or commit SHA.")
        self.root.after(100, self.pump)

    # -- UI helpers (main thread only) --
    def log(self, msg: str) -> None:
        self.out.configure(state="normal")
        self.out.insert(tk.END, msg + "\n")
        self.out.see(tk.END)
        self.out.configure(state="disabled")

    def clear(self) -> None:
        self.out.configure(state="normal")
        self.out.delete("1.0", tk.END)
        self.out.configure(state="disabled")

    def set_busy(self, busy: bool, text: str = "Ready") -> None:
        for b in self.buttons:
            b.configure(state="disabled" if busy else "normal")
        self.stop_btn.configure(state="normal" if busy else "disabled")
        self.status.config(text=text)

    def pump(self) -> None:
        """Apply events queued by worker threads; Tk is not thread-safe."""
        try:
            while True:
                kind, msg = self.events.get_nowait()
                if kind == "log":
                    self.log(msg)
                elif kind == "done":
                    self.log(msg)
                    self.set_busy(False)
        except queue.Empty:
            pass
        self.root.after(100, self.pump)

    # -- actions --
    def stop(self) -> None:
        if self.proc and self.proc.poll() is None:
            self.proc.kill()
            self.log("[Stopped]")

    def send_input(self, _event: object = None) -> None:
        text = self.entry.get()
        if self.proc and self.proc.poll() is None and self.proc.stdin:
            try:
                self.proc.stdin.write(text + "\n")
                self.proc.stdin.flush()
                self.log(f"> {text}")
            except OSError as e:
                self.log(f"[Send error] {e}")
        self.entry.delete(0, tk.END)

    def run_script(self, name: str, desc: str) -> None:
        self.set_busy(True, f"Running: {desc}")
        self.log(f"> {desc}  ({name} @ {REF})")
        threading.Thread(target=self._worker, args=(name,), daemon=True).start()

    def _worker(self, name: str) -> None:
        tmp_path = ""
        try:
            if self.manifest is None:
                self.manifest = fetch_manifest()
            body = fetch_verified(self.manifest, name)
            self.events.put(("log", "SHA-256 verified."))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".ps1", mode="wb") as tmp:
                # BOM so Windows PowerShell 5.1 reads UTF-8 correctly.
                tmp.write(b"\xef\xbb\xbf" + PS_OVERRIDE.encode() + b"\n" + body.removeprefix(b"\xef\xbb\xbf"))
                tmp_path = tmp.name
            exe = shutil.which("powershell.exe") or "powershell.exe"
            self.proc = subprocess.Popen(  # noqa: S603 - fixed argv, verified script
                [exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if self.proc.stdout is None:
                raise OSError("PowerShell stdout pipe unavailable")
            for line in self.proc.stdout:
                self.events.put(("log", line.rstrip()))
            code = self.proc.wait()
            self.events.put(("done", f"[Completed] Exit code: {code}"))
        except (requests.RequestException, ValueError, OSError, KeyError) as e:
            self.events.put(("done", f"[Error] {e}"))
        finally:
            if tmp_path:
                with contextlib.suppress(OSError):
                    os.remove(tmp_path)


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
