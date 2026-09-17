' ==============================================================================
' MRMV Report Content Migration Tool - Windows Launcher
' Double-click this file on Windows to launch without any console window!
' ==============================================================================
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName(WScript.ScriptFullName)

On Error Resume Next
' 1. Fast Launch: Try windowless Python (pythonw.exe)
WshShell.Run "pythonw """ & strPath & "\gui_app.py""", 0, False
If Err.Number <> 0 Then
    Err.Clear
    ' 2. Try pyw launcher
    WshShell.Run "pyw """ & strPath & "\gui_app.py""", 0, False
    If Err.Number <> 0 Then
        Err.Clear
        ' 3. Try standard python
        WshShell.Run "python """ & strPath & "\gui_app.py""", 0, False
        If Err.Number <> 0 Then
            Err.Clear
            ' 4. Try py launcher
            WshShell.Run "py """ & strPath & "\gui_app.py""", 0, False
            If Err.Number <> 0 Then
                MsgBox "Python 3 was not found on this computer." & vbCrLf & vbCrLf & _
                       "Please ensure Python 3.8+ is installed and added to your PATH.", _
                       vbCritical, "MRMV Migration Tool - Error"
            End If
        End If
    End If
End If

Set FSO = Nothing
Set WshShell = Nothing

