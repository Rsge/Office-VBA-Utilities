Attribute VB_Name = "DifferingEntriesFinder"
Attribute VB_Description = "Finds entries that aren't on a different sheet."
'@Folder("Finders")
'@ModuleDescription("Finds entries that aren't on a different sheet.")
Option Explicit

' Integer constants
'@VariableDescription("Sheet index used to confirm content of active sheet.")
Private Const m_confirmSheet As Long = 1
Attribute m_confirmSheet.VB_VarDescription = "Sheet index used to confirm content of active sheet."
'@VariableDescription("With which row to start on active sheet.")
Private Const m_startingRowAct As Long = 1
Attribute m_startingRowAct.VB_VarDescription = "With which row to start on active sheet."
'@VariableDescription("With which row to start on confirm sheet.")
Private Const m_startingRowFirst As Long = 1
Attribute m_startingRowFirst.VB_VarDescription = "With which row to start on confirm sheet."
'@VariableDescription("Which column to search in active sheet.")
Private Const m_searchColumnAct As Long = 1
Attribute m_searchColumnAct.VB_VarDescription = "Which column to search in active sheet."
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
'@Description("Finds entries in active sheet's specified column that aren't in first sheet's specified column.")
Public Sub FindDifferingEntries()
Attribute FindDifferingEntries.VB_Description = "Finds entries in active sheet's specified column that aren't in first sheet's specified column."
    Dim currentActRow As Long
    currentActRow = m_startingRowAct
    Dim currentFirstRow As Long
    Dim currentFoundRows As Long
    currentFoundRows = m_startingRowFirst
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.[_Default](m_confirmSheet)
    Dim check As Boolean
    Dim differing As String
    differing = vbNullString
    Do Until IsEmpty(GetActCellValue(currentActRow, m_searchColumnAct))
        currentFirstRow = currentFoundRows
        check = False
        Do Until IsEmpty(GetCellValue(ws, currentFirstRow, m_searchColumnFirst))
            If GetActCellValue(currentActRow, m_searchColumnAct) = GetCellValue(ws, currentFirstRow, m_searchColumnFirst) Then
                check = True
                currentFoundRows = currentFoundRows + 1
                Exit Do
            End If
            currentFirstRow = currentFirstRow + 1
        Loop
        If Not check Then
            differing = differing & GetActCellValue(currentActRow, m_searchColumnAct) & vbNewLine
        End If
'        If GetActCellValue(i, m_searchColumnActive) <> GetCellValue(ws, i, m_searchColumnFirst) Then
'            differing = differing & GetActCellValue(i, m_searchColumnActive) & vbNewLine
'        End If
        currentActRow = currentActRow + 1
        DoEvents
    Loop
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", differing
        End With
    End With
    MsgBox differing
End Sub
