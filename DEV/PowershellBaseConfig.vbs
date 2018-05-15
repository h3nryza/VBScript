Set oShell = CreateObject("Shell.Application")  
oShell.ShellExecute "%SystemRoot%\syswow64\WindowsPowerShell\v1.0\powershell.exe", "set-executionpolicy bypass", "", "runas", 1 
oShell.ShellExecute "%%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe", "set-executionpolicy bypass", "", "runas", 1 
oShell.ShellExecute "%%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe", "winrm quickconfig /force", "", "runas", 1 
oShell.ShellExecute "%%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe", "Enable-PSRemoting -force", "", "runas", 1 
 
 