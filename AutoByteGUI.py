import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import subprocess, threading, tempfile, requests, os, sys

# Configuration of scripts
GITHUB_SCRIPTS = [
    ("Disk Cleaner",       "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1"),
    ("Enable Files On Demand", "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1"),
    ("Download & Install Package","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadandInstallPackage.ps1"),
    ("Check User Profile",  "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckUserProfileIssue.ps1"),
    ("Dell Bloatware Remover","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1"),
    ("Reset & Install Windows Update","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1"),
    ("Windows System Repair","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1"),
    ("Reset Windows Search DB","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"),
    ("Install MS Projects", "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallMSProjects.ps1"),
    ("Check Drive Space",   "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1"),
    ("Internet Speed Test", "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1"),
    ("Internet Latency Test","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1"),
    ("Monitor Troubleshooter","https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WorkPaperMonitorTroubleShooter.ps1"),
]

# Monkey-patch header to override Read-Host in downloaded scripts
PS_OVERRIDE = r'''
# Redirect Read-Host to stdin pipe
if (Get-Command Read-Host -ErrorAction SilentlyContinue) {
    Remove-Item Function:\Read-Host -ErrorAction SilentlyContinue
}
function Read-Host {
    param([string]$prompt="")
    Write-Host $prompt -NoNewline
    return [Console]::In.ReadLine()
}
'''

current_proc = None

def write_log(txt_widget, message):
    txt_widget.configure(state='normal')
    txt_widget.insert(tk.END, message + '\n')
    txt_widget.see(tk.END)
    txt_widget.configure(state='disabled')

def stop_process():
    global current_proc
    if current_proc and current_proc.poll() is None:
        current_proc.kill()
        write_log(output_box, "[Stopped]")
        reset_ui()

def send_input(event=None):
    global current_proc
    text = input_entry.get()
    if current_proc and current_proc.poll() is None:
        try:
            current_proc.stdin.write(text + '\n')
            current_proc.stdin.flush()
            write_log(output_box, f"> {text}")
        except Exception as e:
            write_log(output_box, f"[Send error] {e}")
    input_entry.delete(0, tk.END)

def reset_ui():
    for btn in script_buttons:
        btn.configure(state='normal')
    stop_btn.configure(state='disabled')
    status_label.config(text="Ready")

def run_script(url, desc):
    global current_proc
    # Disable UI buttons
    for btn in script_buttons:
        btn.configure(state='disabled')
    stop_btn.configure(state='normal')
    status_label.config(text=f"Running: {desc}")
    write_log(output_box, f">▶ {desc}")
    write_log(output_box, f"Downloading from {url}")

    # Download and prefix override
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        script_content = PS_OVERRIDE + '\n' + resp.text
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.ps1', mode='w', encoding='utf-8')
        tmp.write(script_content)
        tmp.close()
        write_log(output_box, f"Saved to {tmp.name}")
    except Exception as e:
        write_log(output_box, f"[Download error] {e}")
        reset_ui()
        return

    # Launch PowerShell with pipes for stdout/stderr/stdin
    current_proc = subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.PIPE, text=True
    )

    # Reader thread
    def reader():
        for line in current_proc.stdout:
            write_log(output_box, line.rstrip())
        code = current_proc.wait()
        os.remove(tmp.name)
        write_log(output_box, f"[Completed] Exit code: {code}")
        reset_ui()
    threading.Thread(target=reader, daemon=True).start()

# Build GUI
root = tk.Tk()
root.title("AutoBytePro GUI")
root.geometry("960x680")

# Button grid (2 rows x 7 columns)
button_frame = tk.Frame(root)
button_frame.pack(fill='x', padx=10, pady=5)

script_buttons = []
cols = 7
for idx, (text, url) in enumerate(GITHUB_SCRIPTS):
    btn = tk.Button(button_frame,
                    text=text,
                    width=16,
                    height=2,
                    wraplength=120,
                    command=lambda u=url, d=text: run_script(u, d))
    row, col = divmod(idx, cols)
    btn.grid(row=row, column=col, padx=4, pady=4)
    script_buttons.append(btn)

# Stop & Clear on right
stop_btn = tk.Button(button_frame, text="■ Stop", width=12, height=2, fg="red", command=stop_process)
stop_btn.grid(row=0, column=cols, padx=(20,4), pady=4)
clear_btn = tk.Button(button_frame, text="🗑 Clear", width=12, height=2,
                      command=lambda: output_box.configure(state='normal') or 
                                        output_box.delete('1.0', tk.END) or
                                        output_box.configure(state='disabled'))
clear_btn.grid(row=1, column=cols, padx=(20,4), pady=4)

# Status bar
status_label = tk.Label(root, text="Ready", anchor='w')
status_label.pack(fill='x', padx=10, pady=(0,5))

# Output console
output_box = ScrolledText(root, state='disabled', font=("Consolas", 10), bg='black', fg='white')
output_box.pack(fill='both', expand=True, padx=10, pady=(0,5))

# Input entry for interactive scripts
input_frame = tk.Frame(root)
input_frame.pack(fill='x', padx=10, pady=(0,10))
tk.Label(input_frame, text="Input:").pack(side='left')
input_entry = tk.Entry(input_frame)
input_entry.pack(side='left', fill='x', expand=True, padx=(5,5))
input_entry.bind("<Return>", send_input)
tk.Button(input_frame, text="Send", command=send_input).pack(side='left')

root.mainloop()
