Attribute VB_Name = "Protection"
Attribute VB_Description = "Module for protection utils."
'@Folder "Attendence time table"
'@ModuleDescription "Module for protection utils."
Option Explicit

'String constants
Private Const ProtectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Private Const ProtectionEnabled As String = "Protection reestablished."


'@EntryPoint
'@Description "Toggle protection status of worksheet."
Public Sub UnProtect()
Attribute UnProtect.VB_Description = "Toggle protection status of worksheet."
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    If WS.ProtectContents = True Then
        WS.UnProtect
        MsgBox (ProtectionLifted)
    Else
        WS.Protect
        MsgBox (ProtectionEnabled)
    End If
End Sub
