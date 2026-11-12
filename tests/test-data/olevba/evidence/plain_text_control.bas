Attribute VB_Name = "Documentation"
' This module is documentation only: every suspicious-looking word is in a
' comment or a string, never in executable code.
' Mentioned for reference: AutoExec, Auto_Open, Shell, Write, Execute,
' CreateObject, WScript.Shell, URLDownloadToFileA, PowerShell.

Sub ExplainBehaviour()
    Dim note As String
    note = "A real attack might call Shell or Write, but this line is only text."
    MsgBox "No Shell, no Write, no Execute statement exists in this module."
    Debug.Print note
End Sub
