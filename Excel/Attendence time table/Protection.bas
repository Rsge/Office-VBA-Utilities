Attribute VB_Name = "Protection"
Attribute VB_Description = "Protection utils."
'@Folder "Attendence time table"
'@ModuleDescription "Protection utils."
Option Explicit

'String constants
'@VariableDescription "Message for lifted worksheet protection."
Private Const ProtectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Attribute ProtectionLifted.VB_VarDescription = "Message for lifted worksheet protection."
'@VariableDescription "Message for enforced worksheet protection."
Private Const ProtectionEnabled As String = "Protection reestablished."
Attribute ProtectionEnabled.VB_VarDescription = "Message for enforced worksheet protection."


'@EntryPoint
'@Description "Toggles protection status of worksheet."
Public Sub UnProtect()
Attribute UnProtect.VB_Description = "Toggles protection status of worksheet."
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
