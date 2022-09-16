Attribute VB_Name = "FixingZipCodes"
'@IgnoreModule
'@Folder "Fixing Zip Codes"
Option Explicit

'@VariableDescription "Path to file for saving names of changed companys."
Private Const m_savePath As String = "C:\Example\Test.txt"
'@VariableDescription "Column with company name."
Private Const m_nameColumnName As String = "Name"
'@VariableDescription "Country to check data of."
Private Const m_countryToCheck As String = "U.S.A."
'@VariableDescription "Name of column with country entries."
Private Const m_countryColumnName As String = "Country"
'@VariableDescription "Column to check a condition in."
Private Const m_checkInColumnName As String = "ZipCode"
'@VariableDescription "Column to transfer data from check column to."
Private Const m_transferToColumnName As String = "City"

'@VariableDescription "The first row with data."
Private Const m_startingRow As Long = 2
'@VariableDescription "How many chars of check cell should be checked."
Private Const m_checkLength As Long = 2


'@EntryPoint
'@Description "Goes through entries of a specific country, checks if the zip code is formatted wrongly and corrects it."
Public Sub FixZipCodes()
    Dim i As Long
    i = m_startingRow
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim table As ListObject
    Set table = ws.ListObjects.[_Default](1)
    Dim tableNameColumn As Long
    tableNameColumn = table.ListColumns.Item(m_nameColumnName).Index
    Dim checkInColumn As Long
    Dim cutPart As String
    Dim transferToColumn As Long
    Dim fixedEntries As Object
    Set fixedEntries = CreateObject("System.Collections.ArrayList")
    
    'Iterate through all columns
    Do Until LenB(ws.Cells.Item(i, tableNameColumn)) = 0
        If ws.Cells.Item(i, table.ListColumns.Item(m_countryColumnName).Index) = m_countryToCheck Then
            checkInColumn = table.ListColumns.Item(m_checkInColumnName).Index
            cutPart = Left$(ws.Cells.Item(i, checkInColumn), m_checkLength)
            If Not IsNumeric(cutPart) Then
                ws.Cells.Item(i, checkInColumn).NumberFormat = "@"
                ws.Cells.Item(i, checkInColumn) = Trim$(Replace(ws.Cells.Item(i, checkInColumn), cutPart, vbNullString))
                transferToColumn = table.ListColumns.Item(m_transferToColumnName).Index
                ws.Cells.Item(i, transferToColumn) = ws.Cells.Item(i, transferToColumn) & ", " & cutPart
                fixedEntries.Add ws.Cells.Item(i, tableNameColumn) & " " & ws.Cells.Item(i, tableNameColumn + 1)
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

