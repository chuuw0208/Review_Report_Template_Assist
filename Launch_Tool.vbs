' ==============================================================================
' MRMV Report Content Migration Tool - Instant Silent Launcher
' Double-click this file to launch the tool without any command prompt window!
' ==============================================================================
Set WshShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' 1. Try launching directly with pythonw (fastest, windowless, <0.2s launch)
On Error Resume Next
WshShell.Run "pythonw """ & strPath & "\gui_app.py""", 0, False
If Err.Number <> 0 Then
    Err.Clear
    ' 2. Try pyw launcher
    WshShell.Run "pyw """ & strPath & "\gui_app.py""", 0, False
    If Err.Number <> 0 Then
        Err.Clear
        ' 3. Fallback to hidden batch launcher
        WshShell.Run "cmd.exe /c """ & strPath & "\run_tool.bat""", 0, False
    End If
End If
Set WshShell = Nothing
