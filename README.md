# Documentation

### Components

#### Task Scheduler Command

The following command is used by the Windows Task Scheduler to run the PowerShell script:

powershell.exe -nologo -windowstyle hidden c:\BackupScripts\GetOpenFiles.ps1

- powershell.exe: Executes the PowerShell script.
- -nologo: Suppresses the PowerShell logo in the output.
- -windowstyle hidden: Runs the script in a hidden window.
- c:\BackupScripts\GetOpenFiles.ps1: Path to the PowerShell script.

#### VBScript

A VBScript is used to execute the PowerShell command:

command = "powershell.exe -nologo -windowstyle hidden c:\BackupScripts\GetOpenFiles.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0

- CreateObject("WScript.Shell"): Creates a WScript.Shell object to run the command.
- shell.Run command,0: Executes the command with no window shown (0).

#### PowerShell Script

The PowerShell script (GetOpenFiles.ps1) generates a report of open files:

Get-SmbOpenFile | Select-Object @('ClientUserName','Path','ShareRelativePath','ClientComputerName') | Format-Table -AutoSize | Out-File -FilePath "E:\Company\General\OPEN_FILE_REPOT.txt" -Width 500

- Get-SmbOpenFile: Retrieves a list of open files on SMB shares.
- Select-Object: Selects specific properties to include in the report:
  - ClientUserName
  - Path
  - ShareRelativePath
  - ClientComputerName
- Format-Table -AutoSize: Formats the output as a table with automatic column sizing.
- Out-File -FilePath "E:\Company\General\OPEN_FILE_REPOT.txt" -Width 500: Outputs the formatted table to a file with a specified width.
