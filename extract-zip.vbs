Dim WshShell, zipPath, psScript, cmd
Set WshShell = CreateObject("WScript.Shell")
zipPath  = WScript.Arguments(0)
psScript = Replace(WScript.ScriptFullName, ".vbs", ".ps1")

cmd = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & psScript & """ -ZipPath """ & zipPath & """"
WshShell.Run cmd, 0, False
