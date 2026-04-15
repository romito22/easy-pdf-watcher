Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
baseFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
psScript = FSO.BuildPath(FSO.BuildPath(baseFolder, "..\main_files"), "setup_pdf_easy_watcher.ps1")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File """ & psScript & """", 1, False
