' ==============================================================================
' MRMV Report Content Migration Tool - Silent Launcher
' Double-click this file to launch the tool without any command prompt window!
' ==============================================================================
Set WshShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' Run run_tool.bat completely hidden (window style 0, no wait)
WshShell.Run "cmd.exe /c """ & strPath & "\run_tool.bat""", 0, False
Set WshShell = Nothing

