# AutoByte - Easy Automate Your Repetitive Troubleshooting Tasks with PowerShell

![AutoByte](https://github.com/deep1ne8/misc/blob/main/AutoByte_Main.png)

AutoByte is a **PowerShell automation tool** that simplifies system tasks, making it easier to manage deployments, configurations, and automation workflows.

## 🤖 AutoByte Agent (new — the headline feature)
A **local, private Windows AI troubleshooting agent**. Talk to it in plain English; it inspects your PC with safe probes, explains issues, and (with your approval) runs fixes from the bundled script library. Runs a small GGUF model on your CPU — no data leaves the machine, no token limit.

```powershell
cd AutoByte-Agent
pip install -r requirements.txt
python setup_model.py      # one-time ~4.7GB model download
python agent.py            # chat with your PC
```
Test the framework without a download: `python agent.py --mock`. Full docs in [`AutoByte-Agent/README.md`](AutoByte-Agent/README.md).

## 🔥 Features
- Local AI agent that troubleshoots, fixes, repairs and updates Windows (propose-and-confirm safety)
- 93 bundled PowerShell scripts (printers, Office/M365, Windows repair, network, Dell/HP, calendar/Exchange, deploy)
- AutoByte CLI — self-discovering script launcher
- AutoByteGUI — Python/Tk front-end

## This is my personal project to improve my PowerShell skills and to automate repetitive troubleshooting tasks.
## Tool in continious development. New features will be added soon!!!!

## 🚀 Installation
You can install AutoByte using **PowerShell**:
```powershell
git clone https://github.com/deep1ne8/misc.git
cd .\misc
& .\AutoByteGUI-Build.ps1
```

Execute AutoByteGUI:
```powershell
& python .\AutoByteGUI.py
```

## 🖥️ AutoByte CLI (new)
A self-discovering launcher for the `Scripts/` library. It reads every
`Scripts/*.ps1`, shows what each does, and runs the one you pick.

```powershell
.\AutoByte-CLI.ps1                # interactive numbered menu
.\AutoByte-CLI.ps1 list           # all scripts + descriptions
.\AutoByte-CLI.ps1 categories     # grouped by inferred category
.\AutoByte-CLI.ps1 show <name>    # description of one script
.\AutoByte-CLI.ps1 run <name>     # run a script by name
```

Example:
```powershell
.\AutoByte-CLI.ps1 run CheckIfUserIsAdmin
```

> Each script runs in its own PowerShell process. Review scripts before
> production use — many require Administrator and modify system state.

## 📖 Documentation
For detailed guides and usage instructions, visit the [AutoByte Wiki](https://github.com/deep1ne8/misc/wiki).

## 📜 Release Notes
Check out the [Release Notes](https://github.com/deep1ne8/misc/releases) for new updates and improvements.

## 📬 Support
- Report issues: [GitHub Issues](https://github.com/deep1ne8/misc/issues)
- Discussions: [Community Discussions](https://github.com/deep1ne8/misc/discussions)
- Contributions: Open-source contributions are welcome! 🎉

[autoByteImage]: https://github.com/deep1ne8/misc/blob/main/AutoByte_Main.png
