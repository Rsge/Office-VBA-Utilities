Attribute VB_Name = "FindingUsedColumns"
'@Folder "ColumnFinding"
Option Explicit

'@VariableDescription "Path to file for saving used column names."
Private Const m_savePath As String = "C:\Example\Test.txt"
Attribute m_savePath.VB_VarDescription = "Path to file for saving used column names."
'@VariableDescription "Ignore columns with names containting this after first space if already present without this part."
Private Const m_ignoreColumnWith As String = "("
Attribute m_ignoreColumnWith.VB_VarDescription = "Ignore columns with names containting this after first space if already present without this part."

'@VariableDescription "The first row with data."
Private Const m_startingRow As Long = 4
Attribute m_startingRow.VB_VarDescription = "The first row with data."
'@VariableDescription "The first column with data."
Private Const m_startingColumn As Long = 1
Attribute m_startingColumn.VB_VarDescription = "The first column with data."
'@VariableDescription "How many rows to check for difference before checking last row."
Private Const m_checkAmount As Long = 3
Attribute m_checkAmount.VB_VarDescription = "How many rows to check for difference before checking last row."

'@EntryPoint
'@Description "Finds columns with at least one difference in content to determine if used and outputs column names to txt."
Public Sub FindUsedColumns()
Attribute FindUsedColumns.VB_Description = "Finds columns with at least one difference in content to determine if used and outputs column names to txt."
    Dim i As Long
    i = m_startingColumn
    Dim j As Long
    j = m_startingRow
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim table As ListObject
    Set table = ws.ListObjects.[_Default](1)
    Dim columnName As String
    Dim columnNameParts() As String
    Dim previousContent As String
    Dim usedColumns As Object
    Set usedColumns = CreateObject("System.Collections.ArrayList")
    Dim lastRow As Long
    
    'Save active workbook to update special cells
    ActiveWorkbook.Save
    
    'Iterate through all columns
    Do Until LenB(ws.Cells.Item(m_startingRow - 1, i)) = 0
        'Sort table by current column's values ascending
        With table.Sort
            .SortFields.Clear
            .SortFields.Add ws.Columns.Item(i), xlSortOnValues, xlAscending
            .Header = xlYes
            .Apply
        End With
        
        'Check if column is worth looking at
        columnName = ws.Cells.Item(m_startingRow - 1, i).Value
        columnNameParts = Split(columnName, " ")
        If UBound(columnNameParts) > 0 Then
            If (InStrB(columnNameParts(UBound(columnNameParts)), m_ignoreColumnWith) = 1 _
                Or IsNumeric(columnNameParts(UBound(columnNameParts)))) _
                And usedColumns.Contains(Replace(columnName, " " & columnNameParts(UBound(columnNameParts)), vbNullString)) _
                Then GoTo Continue
        End If
        'Check if one of the first entries is different from the last one checked.
        previousContent = ws.Cells.Item(j, i)
        Do
            j = j + 1
            If ws.Cells.Item(j, i) <> previousContent Then
                usedColumns.Add columnName
                Exit Do
            End If
        Loop While j < m_checkAmount
        'If the first few entries are the same, check if the last entry differs.
        If Not usedColumns.Contains(columnName) Then
            lastRow = table.Range.SpecialCells(xlCellTypeLastCell).Row - 1
            If ws.Cells.Item(lastRow, i) <> previousContent Then
                usedColumns.Add columnName
            End If
        End If
Continue:
        'Iterate and still be able to look at change
        i = i + 1
        DoEvents
    Loop
    
    'Save results in txt
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim entry As Variant
    Open m_savePath For Output As fileNumber
    For Each entry In usedColumns
        Print #fileNumber, entry
    Next
    Close fileNumber
End Sub
