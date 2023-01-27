Attribute VB_Name = "Cleanup"
Attribute VB_Description = "Cleanup utils."
'@Folder("BasicUtilities")
'@ModuleDescription("Cleanup utils.")
Option Explicit

' String constants
'@VariableDescription("Area in active worksheet which is to clear.")
Private Const m_areaToClear As String = "A1:B2,C1:D2"
Attribute m_areaToClear.VB_VarDescription = "Area in active worksheet which is to clear."
'@VariableDescription("Title of MsgBox to show it contains a warning.")
Private Const m_warningLabel As String = "Warning!"
Attribute m_warningLabel.VB_VarDescription = "Title of MsgBox to show it contains a warning."
'@VariableDescription("Warning about the specified area being deleted if accepted.")
Private Const m_deletionWarning As String = "Everything in this table will be deleted!"
Attribute m_deletionWarning.VB_VarDescription = "Warning about the specified area being deleted if accepted."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Deletes all cells' content in defined area.")
Public Sub Clear()
Attribute Clear.VB_Description = "Deletes all cells' content in defined area."
    Dim Check As Byte
    Check = MsgBox(m_deletionWarning, vbOKCancel + vbExclamation, m_warningLabel)
    If Check = 1 Then
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.ActiveSheet
        ws.UnProtect
        Dim CellsToDelete As Range
        Set CellsToDelete = ActiveSheet.Range(m_areaToClear)
        CellsToDelete.ClearContents
        ws.Protect
    End If
End Sub
