Attribute VB_Name = "EntryEqualizer"
Attribute VB_Description = "Finds and deletes entries that aren't on a different sheet."
'@Folder("Finders")
'@ModuleDescription("Finds and deletes entries that aren't on a different sheet.")
Option Explicit

' String constants
Private Const m_actColumnQuestion As String = "Which column contains the comparison data?" & vbNewLine _
                                            & "(A, B, ..)"
Private Const m_actColumnInputError As String = "No column found with this label." & vbNewLine _
                                              & "Please try again."

' Integer constants
'@VariableDescription("Sheet index with entries used as a basis for active sheet.")
Private Const m_confirmSheet As Long = 1
Attribute m_confirmSheet.VB_VarDescription = "Sheet index with entries used as a basis for active sheet."
'@VariableDescription("With which row to start on active sheet.")
Private Const m_startingRowAct As Long = 4
Attribute m_startingRowAct.VB_VarDescription = "With which row to start on active sheet."
'@VariableDescription("With which row to start on confirm sheet.")
Private Const m_startingRowFirst As Long = 1
Attribute m_startingRowFirst.VB_VarDescription = "With which row to start on confirm sheet."
'@VariableDescription("Which column to search in confirm sheet.")
Private Const m_searchColumnFirst As Long = 1
Attribute m_searchColumnFirst.VB_VarDescription = "Which column to search in confirm sheet."


' ————————————————————————————————————————————————————— '

'@Description("Tests if a string is empty.")
Private Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

'@Description("Gets the value of a cell on a worksheet at a position.")
Private Function GetCellValue(ByVal ws As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetCellValue.VB_Description = "Gets the value of a cell on a worksheet at a position."
    GetCellValue = ws.Cells.Item(row_, column_).Value
End Function

'@Description("Gets the value of a cell on the active worksheet at a position.")
Private Function GetActCellValue(ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetActCellValue.VB_Description = "Gets the value of a cell on the active worksheet at a position."
    GetActCellValue = GetCellValue(ActiveSheet, row_, column_)
End Function

'@EntryPoint
'@Description("Deletes all entries in active sheet's given column that aren't in first sheet's specified column.")
Public Sub EqualizeEntries()
Attribute EqualizeEntries.VB_Description = "Deletes all entries in active sheet's given column that aren't in first sheet's specified column."
    ' Ask for relevant column on active sheet.
    Dim searchColumnActAnswer As String
    Dim searchColumnAct As Long
    If False Then
Retry:
        MsgBox m_actColumnInputError, vbExclamation
    End If
    searchColumnActAnswer = InputBox(m_actColumnQuestion)
    If IsEmpty(searchColumnActAnswer) Then Exit Sub
    On Error GoTo Retry
    searchColumnAct = ActiveSheet.Columns(searchColumnActAnswer).Column
    On Error GoTo 0
    ' Prepare variables.
    Dim currentRow As Long
    currentRow = m_startingRowFirst
    Dim confirmWS As Worksheet
    Set confirmWS = ActiveWorkbook.Sheets.[_Default](m_confirmSheet)
    Dim entries As Object
    Set entries = CreateObject("System.Collections.ArrayList")
    ' Get entries from confirm sheet.
    Do Until IsEmpty(GetCellValue(confirmWS, currentRow, m_searchColumnFirst))
        entries.Add GetCellValue(confirmWS, currentRow, m_searchColumnFirst)
        currentRow = currentRow + 1
        DoEvents
    Loop
    ' Delete rows of entries on active sheet which are not on confirm sheet.
    currentRow = m_startingRowAct
    Do Until IsEmpty(GetActCellValue(currentRow, searchColumnAct))
        If Not entries.Contains(GetActCellValue(currentRow, searchColumnAct)) Then
            ActiveSheet.Rows(currentRow).Delete
        Else
            currentRow = currentRow + 1
        End If
        ActiveSheet.Rows(currentRow).Select
        DoEvents
    Loop
End Sub
