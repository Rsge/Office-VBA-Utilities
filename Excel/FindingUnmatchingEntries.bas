Attribute VB_Name = "FindingUnmatchingEntries"
'@Folder "Unmatch Finding"
Option Explicit

'Constants
'@VariableDescription "Sheet index used to confirm content of active sheet."
Private Const m_confirmSheet As Long = 1
Attribute m_confirmSheet.VB_VarDescription = "Sheet index used to confirm content of active sheet."
'@VariableDescription "With which row to start on active sheet."
Private Const m_startingRowActive As Long = 1
Attribute m_startingRowActive.VB_VarDescription = "With which row to start on active sheet."
'@VariableDescription "With which row to start on confirm sheet."
Private Const m_startingRowFirst As Long = 1
Attribute m_startingRowFirst.VB_VarDescription = "With which row to start on confirm sheet."
'@VariableDescription "Which column to search in active sheet."
Private Const m_searchColumnActive As Long = 1
Attribute m_searchColumnActive.VB_VarDescription = "Which column to search in active sheet."
'@VariableDescription "Which column to search in confirm sheet."
Private Const m_searchColumnFirst As Long = 1
Attribute m_searchColumnFirst.VB_VarDescription = "Which column to search in confirm sheet."

'@EntryPoint
'@Description "Finds entries in active sheet's specified column that aren't in first sheet's specified column."
Public Sub FindNotIn()
Attribute FindNotIn.VB_Description = "Finds entries in active sheet's specified column that aren't in first sheet's specified column."
    Dim i As Long
    i = m_startingRowActive
    Dim j As Long
    Dim z As Long
    z = m_startingRowFirst
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.[_Default](m_confirmSheet)
    Dim check As Boolean
    Dim unmatched As String
    unmatched = vbNullString
    Do Until LenB(ActiveSheet.Cells.Item(i, m_searchColumnActive)) = 0
        j = z
        check = False
        Do Until LenB(ws.Cells.Item(j, m_searchColumnFirst)) = 0
            If ActiveSheet.Cells.Item(i, m_searchColumnActive) = ws.Cells.Item(j, m_searchColumnFirst) Then
                check = True
                z = z + 1
                Exit Do
            End If
            j = j + 1
        Loop
        If Not check Then
            unmatched = unmatched & ActiveSheet.Cells.Item(i, m_searchColumnActive) & vbNewLine
        End If
'        If ActiveSheet.Cells.Item(i, m_searchColumnActive) <> ws.Cells.Item(i, m_searchColumnFirst) Then
'            unmatched = unmatched & ActiveSheet.Cells.Item(i, m_searchColumnActive) & vbNewLine
'        End If
        i = i + 1
        DoEvents
    Loop
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", unmatched
        End With
    End With
    MsgBox unmatched
End Sub
