command = "powershell.exe -nologo -windowstyle hidden c:\BackupScripts\GetOpenFiles.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0