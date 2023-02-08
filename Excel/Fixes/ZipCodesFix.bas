Attribute VB_Name = "ZipCodesFix"
Attribute VB_Description = "Fixes wrongly formatted zip codes."
'@Folder("Fixes")
'@ModuleDescription("Fixes wrongly formatted zip codes.")
Option Explicit

'@VariableDescription("Path to file for saving names of changed companys.")
Private Const m_savePath As String = "C:\Example\Test.txt"
Attribute m_savePath.VB_VarDescription = "Path to file for saving names of changed companys."
'@VariableDescription("Country to check data of.")
Private Const m_countryToCheck As String = "CANADA"
Attribute m_countryToCheck.VB_VarDescription = "Country to check data of."
'@VariableDescription("Column with company name.")
Private Const m_nameColumnName As String = "Name"
Attribute m_nameColumnName.VB_VarDescription = "Column with company name."
'@VariableDescription("Name of column with country entries.")
Private Const m_countryColumnName As String = "Country"
Attribute m_countryColumnName.VB_VarDescription = "Name of column with country entries."
'@VariableDescription("Column to check a condition in.")
Private Const m_checkInColumnName As String = "ZipCode"
Attribute m_checkInColumnName.VB_VarDescription = "Column to check a condition in."
'@VariableDescription("Column to transfer data from check column to.")
Private Const m_transferToColumnName As String = "City"
Attribute m_transferToColumnName.VB_VarDescription = "Column to transfer data from check column to."
'@VariableDescription("Column to check and transfer data internally.")
Private Const m_checkAndTransferColumnName As String = "POBox"
Attribute m_checkAndTransferColumnName.VB_VarDescription = "Column to check and transfer data internally."
'@VariableDescription("Symbol separating data from irrelevant part in internal transfer column.")
Private Const m_separatingSymbol As String = ", "
Attribute m_separatingSymbol.VB_VarDescription = "Symbol separating data from irrelevant part in internal transfer column."

'@VariableDescription("The first row with data.")
Private Const m_startingRow As Long = 2
Attribute m_startingRow.VB_VarDescription = "The first row with data."
'@VariableDescription("How many chars of check cell should be checked.")
Private Const m_checkLength As Long = 2
Attribute m_checkLength.VB_VarDescription = "How many chars of check cell should be checked."

' ————————————————————————————————————————————————————— '


'@Description("Gets the cell on a worksheet at a position.")
Private Function GetCell(ByVal sheet As Worksheet, ByVal row As Long, ByVal column As Long) As Range
Attribute GetCell.VB_Description = "Gets the cell on a worksheet at a position."
    Set GetCell = sheet.Cells.Item(row, column)
End Function

'@Description("Gets the value of a cell on a worksheet at a position.")
Private Function GetCellValue(ByVal sheet As Worksheet, ByVal row As Long, ByVal column As Long) As Variant
Attribute GetCellValue.VB_Description = "Gets the value of a cell on a worksheet at a position."
    GetCellValue = GetCell(sheet, row, column).Value
End Function

'@Description("Gets the column in a table from it's name.")
Private Function GetTableColumnFromName(ByVal table As ListObject, ByVal columnName As String) As Long
Attribute GetTableColumnFromName.VB_Description = "Gets the column in a table from it's name."
    GetTableColumnFromName = table.ListColumns.Item(columnName).Index
End Function

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Iterates through entries of a specific country, checks if the zip code is formatted wrongly and corrects it.")
Public Sub FixZipCodes()
Attribute FixZipCodes.VB_Description = "Iterates through entries of a specific country, checks if the zip code is formatted wrongly and corrects it."
    Dim i As Long
    i = m_startingRow
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim table As ListObject
    Set table = ws.ListObjects.[_Default](1)
    Dim tableNameColumn As Long
    tableNameColumn = GetTableColumnFromName(table, m_nameColumnName)
    Dim name As String
    Dim checkInColumn As Long
    Dim checkInCell As Range
    Dim cutPart As String
    Dim transferToColumn As Long
    Dim transferToCell As Range
    Dim fixedEntries As Object
    Set fixedEntries = CreateObject("System.Collections.ArrayList")
    Dim parts() As String
    
    ' Iterate through all columns.
    Do Until LenB(GetCellValue(ws, i, tableNameColumn)) = 0
        If GetCellValue(ws, i, GetTableColumnFromName(table, m_countryColumnName)) = m_countryToCheck Then
            ' Zip code and city field
            checkInColumn = GetTableColumnFromName(table, m_checkInColumnName)
            Set checkInCell = GetCell(ws, i, checkInColumn)
            cutPart = Left$(checkInCell.Value, m_checkLength)
            name = GetCellValue(ws, i, tableNameColumn) & " " & GetCellValue(ws, i, tableNameColumn + 1)
            If Not IsNumeric(Right$(cutPart, 1)) Then
                checkInCell.NumberFormat = "@"
                checkInCell.Value = Trim$(Replace(checkInCell.Value, cutPart, vbNullString))
                transferToColumn = table.ListColumns.Item(m_transferToColumnName).Index
                Set transferToCell = ws.Cells.Item(i, transferToColumn)
                transferToCell.Value = transferToCell.Value & m_separatingSymbol & cutPart
                fixedEntries.Add name
            End If
            ' Post box field
            checkInColumn = GetTableColumnFromName(table, m_checkAndTransferColumnName)
            Set checkInCell = ws.Cells.Item(i, checkInColumn)
            parts = Split(checkInCell.Value, m_separatingSymbol)
            If UBound(parts) > 0 Then
                cutPart = Left$(parts(1), m_checkLength)
                If Not IsNumeric(Right$(cutPart, 1)) Then
                    checkInCell.Value = Replace(checkInCell.Value, cutPart & " ", vbNullString) _
                                        & m_separatingSymbol & cutPart
                    If Not fixedEntries.Contains(name) Then
                        fixedEntries.Add name
                    End If
                End If
            End If
        End If
        i = i + 1
    Loop
    
    ' Save results in txt.
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim entry As Variant
    Open m_savePath For Output As fileNumber
    For Each entry In fixedEntries
        Print #fileNumber, entry
    Next
    Close fileNumber
End Sub
