Stop-Service spooler
Remove-Item -Path "HKCU:\Printers\Connections\*" -Recurse -Force
Remove-Item -Path "C:\Windows\System32\spool\SERVERS\CVEFS04\PRINTERS\*" -Recurse -Force -ErrorAction SilentlyContinue
Start-Service spooler
gpupdate /force
