Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
baseFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
psScript = FSO.BuildPath(FSO.BuildPath(baseFolder, "..\.bat"), "main_pdf_easy_watcher.ps1")
WshShell.Run "powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psScript & """", 0, False