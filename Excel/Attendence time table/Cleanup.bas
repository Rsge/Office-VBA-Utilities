Attribute VB_Name = "Cleanup"
Option Explicit
Sub Clear()
    Dim Warning As String
    Warning = "Everything in this table will be deleted!"
    Dim Check As Byte
    Check = MsgBox(Warning, vbOKCancel + vbExclamation, "Warning!")
    If Check = 1 Then
        Dim WS As Worksheet
        Set WS = ActiveWorkbook.ActiveSheet
        WS.UnProtect
        Dim CellsToDelete As Range
        Set CellsToDelete = Range("C9:AG52,C55:AG72,C75:AG97,C100:AG119,C122:AG137,C140:AG142,C146:AG149,C152:AG152")
        CellsToDelete.ClearContents
        WS.Protect
    End If
End Sub

