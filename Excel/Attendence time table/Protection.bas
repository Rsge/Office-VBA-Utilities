Attribute VB_Name = "Protection"
Option Explicit
Public Sub UnProtect()
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    Dim Whatever
    If WS.ProtectContents = True Then
        WS.UnProtect
        Whatever = MsgBox("Protection lifted." & vbCrLf & "Changes now possible.")
    Else
        WS.Protect
        Whatever = MsgBox("Protection reestablished.")
    End If
End Sub
