# AutoByte - Easy Automate Your Workflow with PowerShell

![AutoByte](https://github.com/deep1ne8/misc/blob/main/AutoByte_Main.png)

AutoByte is a **PowerShell automation tool** that simplifies system tasks, making it easier to manage deployments, configurations, and automation workflows.

## 🔥 Features
- Automates repetitive system tasks
- Lightweight and fast execution
- Seamless integration with PowerShell
- Open-source and customizable

## 🚀 Installation (via Chocolatey)
You can install AutoByte using **Chocolatey**:
```powershell
choco install autobyte -y
```

To uninstall AutoByte:
```powershell
choco uninstall autobyte -y
```

## 📖 Documentation
For detailed guides and usage instructions, visit the [AutoByte Wiki](https://github.com/deep1ne8/misc/wiki).

## 📜 Release Notes
Check out the [Release Notes](https://github.com/deep1ne8/misc/releases) for new updates and improvements.

## 📬 Support
- Report issues: [GitHub Issues](https://github.com/deep1ne8/misc/issues)
- Discussions: [Community Discussions](https://github.com/deep1ne8/misc/discussions)
- Contributions: Open-source contributions are welcome! 🎉

[autoByteImage]: https://github.com/deep1ne8/misc/blob/main/AutoByte_Main.png
## ImmyBotClone

`ImmyBotClone.ps1` provides a simple example of how you can build a lightweight automation framework similar to ImmyBot. It maintains a device inventory (`Devices.json`) and allows you to execute scripts from the repository on remote machines using PowerShell remoting.

### Prerequisites
- PowerShell 7 or later
- Remote systems should have PowerShell remoting enabled (`Enable-PSRemoting -Force`)

On Debian-based systems you can install PowerShell with:

```bash
sudo apt-get update
sudo apt-get install -y powershell
```

Run the script with PowerShell:

```powershell
./ImmyBotClone.ps1
```

You can then add devices, list them and execute scripts from the `Scripts` folder on those devices.

