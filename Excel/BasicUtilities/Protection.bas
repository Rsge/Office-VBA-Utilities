Attribute VB_Name = "Protection"
Attribute VB_Description = "Protection utils."
'@Folder("BasicUtilities")
'@ModuleDescription("Protection utils.")
Option Explicit

' String constants
'@VariableDescription("Message for lifted worksheet protection.")
Private Const m_protectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Attribute m_protectionLifted.VB_VarDescription = "Message for lifted worksheet protection."
'@VariableDescription("Message for enforced worksheet protection.")
Private Const m_protectionEnabled As String = "Protection reestablished."
Attribute m_protectionEnabled.VB_VarDescription = "Message for enforced worksheet protection."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Toggles protection status of worksheet.")
Public Sub UnProtect()
Attribute UnProtect.VB_Description = "Toggles protection status of worksheet."
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    If ws.ProtectContents Then
        ws.UnProtect
        MsgBox (m_protectionLifted)
    Else
        ws.Protect
        MsgBox (m_protectionEnabled)
    End If
End Sub
