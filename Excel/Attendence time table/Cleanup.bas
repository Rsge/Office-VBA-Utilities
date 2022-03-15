Attribute VB_Name = "Cleanup"
Attribute VB_Description = "Module for cleanup utils."
'@Folder "Attendence time table"
'@ModuleDescription "Module for cleanup utils."
Option Explicit

'String constants
Private Const AreaToClear As String = "A1:B2,C1:D2"
Private Const WarningLabel As String = "Warning!"
Private Const DeletionWarning As String = "Everything in this table will be deleted!"


'@EntryPoint
'@Description "Delete all cells' content in defined area."
Public Sub Clear()
Attribute Clear.VB_Description = "Delete all cells' content in defined area."
    Dim Check As Byte
    Check = MsgBox(DeletionWarning, vbOKCancel + vbExclamation, WarningLabel)
    If Check = 1 Then
        Dim WS As Worksheet
        Set WS = ActiveWorkbook.ActiveSheet
        WS.UnProtect
        Dim CellsToDelete As Range
        Set CellsToDelete = ActiveSheet.Range(AreaToClear)
        CellsToDelete.ClearContents
        WS.Protect
    End If
End Sub
