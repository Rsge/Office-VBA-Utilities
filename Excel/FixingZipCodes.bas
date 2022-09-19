Attribute VB_Name = "FixingZipCodes"
'@Folder "Fixing Zip Codes"
Option Explicit

'@VariableDescription "Path to file for saving names of changed companys."
Private Const m_savePath As String = "C:\Example\Test.txt"
Attribute m_savePath.VB_VarDescription = "Path to file for saving names of changed companys."
'@VariableDescription "Column with company name."
Private Const m_nameColumnName As String = "Name"
Attribute m_nameColumnName.VB_VarDescription = "Column with company name."
'@VariableDescription "Country to check data of."
Private Const m_countryToCheck As String = "U.S.A."
Attribute m_countryToCheck.VB_VarDescription = "Country to check data of."
'@VariableDescription "Name of column with country entries."
Private Const m_countryColumnName As String = "Country"
Attribute m_countryColumnName.VB_VarDescription = "Name of column with country entries."
'@VariableDescription "Column to check a condition in."
Private Const m_checkInColumnName As String = "ZipCode"
Attribute m_checkInColumnName.VB_VarDescription = "Column to check a condition in."
'@VariableDescription "Column to transfer data from check column to."
Private Const m_transferToColumnName As String = "City"
Attribute m_transferToColumnName.VB_VarDescription = "Column to transfer data from check column to."
'@VariableDescription "Column to check and transfer data internally."
Private Const m_checkAndTransferColumnName As String = "P.O.Box"
Attribute m_checkAndTransferColumnName.VB_VarDescription = "Column to check and transfer data internally."
'@VariableDescription "Symbol separating data from irrelevant part in internal transfer column."
Private Const m_separatingSymbol As String = ", "
Attribute m_separatingSymbol.VB_VarDescription = "Symbol separating data from irrelevant part in internal transfer column."

'@VariableDescription "The first row with data."
Private Const m_startingRow As Long = 2
Attribute m_startingRow.VB_VarDescription = "The first row with data."
'@VariableDescription "How many chars of check cell should be checked."
Private Const m_checkLength As Long = 2
Attribute m_checkLength.VB_VarDescription = "How many chars of check cell should be checked."


'@EntryPoint
'@Description "Goes through entries of a specific country, checks if the zip code is formatted wrongly and corrects it."
Public Sub FixZipCodes()
Attribute FixZipCodes.VB_Description = "Goes through entries of a specific country, checks if the zip code is formatted wrongly and corrects it."
    Dim i As Long
    i = m_startingRow
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim table As ListObject
    Set table = ws.ListObjects.[_Default](1)
    Dim tableNameColumn As Long
    tableNameColumn = table.ListColumns.Item(m_nameColumnName).Index
    Dim name As String
    Dim checkInColumn As Long
    Dim cutPart As String
    Dim transferToColumn As Long
    Dim fixedEntries As Object
    Set fixedEntries = CreateObject("System.Collections.ArrayList")
    Dim parts() As String
    
    'Iterate through all columns
    Do Until LenB(ws.Cells.Item(i, tableNameColumn)) = 0
        If ws.Cells.Item(i, table.ListColumns.Item(m_countryColumnName).Index) = m_countryToCheck Then
            'Zip code and city field
            checkInColumn = table.ListColumns.Item(m_checkInColumnName).Index
            cutPart = Left$(ws.Cells.Item(i, checkInColumn), m_checkLength)
            name = ws.Cells.Item(i, tableNameColumn) & " " & ws.Cells.Item(i, tableNameColumn + 1)
            If Not IsNumeric(cutPart) Then
                ws.Cells.Item(i, checkInColumn).NumberFormat = "@"
                ws.Cells.Item(i, checkInColumn) = Trim$(Replace(ws.Cells.Item(i, checkInColumn), cutPart, vbNullString))
                transferToColumn = table.ListColumns.Item(m_transferToColumnName).Index
                ws.Cells.Item(i, transferToColumn) = ws.Cells.Item(i, transferToColumn) & m_separatingSymbol & cutPart
                fixedEntries.Add name
            End If
            'Post box field
            checkInColumn = table.ListColumns.Item(m_checkAndTransferColumnName).Index
            parts = Split(ws.Cells.Item(i, checkInColumn), m_separatingSymbol)
            If UBound(parts) > 0 Then
                cutPart = Left$(parts(1), m_checkLength)
                If Not IsNumeric(cutPart) Then
                    ws.Cells.Item(i, checkInColumn) = Replace(ws.Cells.Item(i, checkInColumn), cutPart & " ", vbNullString) _
                                                        & m_separatingSymbol & cutPart
                    If Not fixedEntries.Contains(name) Then
                        fixedEntries.Add name
                    End If
                End If
            End If
        End If
        i = i + 1
    Loop
    
    'Save results in txt
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim entry As Variant
    Open m_savePath For Output As fileNumber
    For Each entry In fixedEntries
        Print #fileNumber, entry
    Next
    Close fileNumber
End Sub
