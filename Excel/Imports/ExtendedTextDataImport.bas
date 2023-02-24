Attribute VB_Name = "ExtendedTextDataImport"
Attribute VB_Description = "Imports extended text data for items from another Excel workbook."
'@Folder("Imports")
'@ModuleDescription("Imports extended text data for items from another Excel workbook.")
Option Explicit

' String constants
'@VariableDescription("Path to workbook for data import.")
Private Const m_wbPath As String = "C:\Example\Test.xlsm"
Attribute m_wbPath.VB_VarDescription = "Path to workbook for data import."
'@VariableDescription("Name of sheet to import from.")
Private Const m_importSheetName As String = "Example"
Attribute m_importSheetName.VB_VarDescription = "Name of sheet to import from."
'@VariableDescription("Name of sheet with header data to import into.")
Private Const m_headerSheetName As String = "ExtendedTextHeader"
Attribute m_headerSheetName.VB_VarDescription = "Name of sheet with header data to import into."
'@VariableDescription("Name of sheet with lines data to import into.")
Private Const m_lineSheetName As String = "ExtendedTextLine"
Attribute m_lineSheetName.VB_VarDescription = "Name of sheet with lines data to import into."
'@VariableDescription("Localized label of True state.")
Private Const m_trueLabel As String = "Yes"
Attribute m_trueLabel.VB_VarDescription = "Localized label of True state."
'@VariableDescription("Localized label of False state.")
Private Const m_falseLabel As String = "No"
Attribute m_falseLabel.VB_VarDescription = "Localized label of False state."
'@VariableDescription("Label for german language line.")
Private Const m_nativeLangCode As String = "EN"
Attribute m_nativeLangCode.VB_VarDescription = "Label for german language line."
'@VariableDescription("Label for english language line.")
Private Const m_translatedLangCode As String = "DE"
Attribute m_translatedLangCode.VB_VarDescription = "Label for english language line."

' Integer constants
'@VariableDescription("Row in which data starts in import sheet.")
Private Const m_importStartingRow As Long = 2
Attribute m_importStartingRow.VB_VarDescription = "Row in which data starts in import sheet."
'@VariableDescription("Row in which data starts in data sheets.")
Private Const m_dataStartingRow As Long = 4
Attribute m_dataStartingRow.VB_VarDescription = "Row in which data starts in data sheets."
'@VariableDescription("Column of item number in import sheet.")
Private Const m_importItemColumn As Long = 1
Attribute m_importItemColumn.VB_VarDescription = "Column of item number in import sheet."
'@VariableDescription("Column of native phrase in import sheet.")
Private Const m_importTranslatedColumn As Long = 4
Attribute m_importTranslatedColumn.VB_VarDescription = "Column of native phrase in import sheet."
'@VariableDescription("Column of translated phrase in import sheet.")
Private Const m_importNativeColumn As Long = 3
Attribute m_importNativeColumn.VB_VarDescription = "Column of translated phrase in import sheet."
'@VariableDescription("Column of item number in data sheets.")
Private Const m_dataItemColumn As Long = 2
Attribute m_dataItemColumn.VB_VarDescription = "Column of item number in data sheets."
'@VariableDescription("Column of language code in data sheets.")
Private Const m_dataLangCodeColumn As Long = 3
Attribute m_dataLangCodeColumn.VB_VarDescription = "Column of language code in data sheets."
'@VariableDescription("Column of text id in data sheets.")
Private Const m_dataTxtNumColumn As Long = 4
Attribute m_dataTxtNumColumn.VB_VarDescription = "Column of text id in data sheets."
'@VariableDescription("Column of start date in header sheet.")
Private Const m_headerStartDateColumn As Long = 5
Attribute m_headerStartDateColumn.VB_VarDescription = "Column of start date in header sheet."
'@VariableDescription("Column of end date in header sheet.")
Private Const m_headerEndDateColumn As Long = 6
Attribute m_headerEndDateColumn.VB_VarDescription = "Column of end date in header sheet."
'@VariableDescription("Column of all lang bool in header sheet.")
Private Const m_headerAllLangColumn As Long = 7
Attribute m_headerAllLangColumn.VB_VarDescription = "Column of all lang bool in header sheet."
'@VariableDescription("Column of line number in lines sheet.")
Private Const m_lineLineNumColumn As Long = 5
Attribute m_lineLineNumColumn.VB_VarDescription = "Column of line number in lines sheet."
'@VariableDescription("Column of text in lines sheet.")
Private Const m_lineTextColumn As Long = 6
Attribute m_lineTextColumn.VB_VarDescription = "Column of text in lines sheet."
'@VariableDescription("The base (= 1) line number in data sheet.")
Private Const m_baseLineNum As Long = 10000
Attribute m_baseLineNum.VB_VarDescription = "The base (= 1) line number in data sheet."

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

'@Description("Inserts a new row at given position in table.")
Private Sub CreateNewRow(ByVal ws As Worksheet, ByVal row As Long)
Attribute CreateNewRow.VB_Description = "Inserts a new row at given position in table."
    ws.Rows.Item(row).Insert
    ws.Rows.Item(row + 1).Copy ws.Rows.Item(row)
End Sub

'@Description("Sets text header cells' values.")
Private Sub SetHeaderCells(ByVal ws As Worksheet, ByVal row As Long)
Attribute SetHeaderCells.VB_Description = "Sets text header cells' values."
    ws.Cells.Item(row, m_headerAllLangColumn).Value = m_falseLabel
    ws.Cells.Item(row, m_dataTxtNumColumn).Value = 0
    Dim langCodeCell As Range
    Set langCodeCell = ws.Cells.Item(row, m_dataLangCodeColumn)
    If ws.Cells.Item(row + 1, m_dataLangCodeColumn).Value = m_translatedLangCode Then
        langCodeCell.Value = m_nativeLangCode
    Else
        langCodeCell.Value = m_translatedLangCode
    End If
    Dim headerEndDateCell As Range
    Set headerEndDateCell = ws.Cells.Item(row, m_headerEndDateColumn)
    If LenB(headerEndDateCell.Value) > 0 Then
        ws.Cells.Item(row, m_headerStartDateColumn).Value = vbNullString
        headerEndDateCell.Value = vbNullString
    End If
End Sub

'@Description("Adds an amount to a cell on a worksheet at a position.")
Private Sub AddToCellValue(ByVal sheet As Worksheet, ByVal row As Long, ByVal column As Long, ByVal addition As Long)
Attribute AddToCellValue.VB_Description = "Adds an amount to a cell on a worksheet at a position."
    GetCell(sheet, row, column).Value = GetCellValue(sheet, row, column) + addition
End Sub

'@Description("Adds localized line text.")
Private Sub AddNewLocalizedTextLine(ByVal importWS As Worksheet, ByVal lineWS As Worksheet, ByVal importRow As Long, ByVal lineRow As Long, ByVal itemNum As String, ByVal langCode As String, ByVal importColumn As Long)
Attribute AddNewLocalizedTextLine.VB_Description = "Adds localized line text."
    CreateNewRow lineWS, lineRow - 1
    lineWS.Cells.Item(lineRow, m_dataItemColumn).Value = itemNum
    lineWS.Cells.Item(lineRow, m_dataLangCodeColumn).Value = langCode
    lineWS.Cells.Item(lineRow, m_dataTxtNumColumn).Value = 1
    lineWS.Cells.Item(lineRow, m_lineLineNumColumn).Value = m_baseLineNum
    lineWS.Cells.Item(lineRow, m_lineTextColumn).Value = importWS.Cells.Item(importRow, importColumn)
End Sub

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Imports data for items from another Excel workbook.")
Public Sub ImportExtendedText()
Attribute ImportExtendedText.VB_Description = "Imports data for items from another Excel workbook."
    Dim importWB As Workbook
    Set importWB = Workbooks.Open(m_wbPath)
    Dim importWS As Worksheet
    Set importWS = importWB.Sheets.[_Default](m_importSheetName)
    Dim headerWS As Worksheet
    Set headerWS = ThisWorkbook.Sheets.[_Default](m_headerSheetName)
    Dim lineWS As Worksheet
    Set lineWS = ThisWorkbook.Sheets.[_Default](m_lineSheetName)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim itemNum As String
    Dim langCode As String
    Dim found(1) As Boolean
    Dim dataLangCodeCell As Range
    Dim lineTextCell As Range
    
    ' Delete start date if no end date exists.
    i = m_dataStartingRow
    Do Until LenB(GetCellValue(headerWS, i, m_dataItemColumn)) = 0
        If LenB(GetCellValue(headerWS, i, m_headerStartDateColumn)) > 0 _
        And LenB(GetCellValue(headerWS, i, m_headerEndDateColumn)) = 0 Then
            GetCell(headerWS, i, m_headerStartDateColumn).Value = vbNullString
        End If
        i = i + 1
    Loop

    ' Go through import data item by item.
    i = m_importStartingRow
    Do Until LenB(GetCellValue(importWS, i, m_importItemColumn)) = 0
        itemNum = GetCellValue(importWS, i, m_importItemColumn)
        ' Find item in header data.
        j = m_dataStartingRow
        Do Until LenB(GetCellValue(headerWS, j, m_dataItemColumn)) = 0
            ' If item is found, process it and exit header loop.
            If itemNum = GetCellValue(headerWS, j, m_dataItemColumn) Then
                ' If next row isn't same item...
                If GetCellValue(headerWS, j + 1, m_dataItemColumn) <> itemNum Then
                    ' If for all langs, change to single lang and add copied row for other lang.
                    ' Else copy row and change to other lang afterwards.
                    If GetCellValue(headerWS, j, m_headerAllLangColumn) = m_trueLabel Then
                        SetHeaderCells headerWS, j
                        CreateNewRow headerWS, j
                        SetHeaderCells headerWS, j
                    Else
                        CreateNewRow headerWS, j
                        SetHeaderCells headerWS, j
                    End If
                ElseIf GetCellValue(headerWS, j, m_headerAllLangColumn) = m_trueLabel 
                And GetCellValue(headerWS, j + 2, m_dataItemColumn) <> itemNum Then
                    SetHeaderCells headerWS, j
                End If
                found(0) = True
                Exit Do
            End If
            j = j + 1
        Loop
        If Not found(0) Then
            CreateNewRow headerWS, j - 1
            SetHeaderCells headerWS, j
            GetCell(headerWS, j, m_dataItemColumn).Value = itemNum
            CreateNewRow headerWS, j
            SetHeaderCells headerWS, j
        End If
        found(0) = False
        ' Find item in line data.
        j = m_dataStartingRow
        Do Until LenB(GetCellValue(lineWS, j, m_dataItemColumn)) = 0
            ' If item is found, process it and exit line loop.
            If itemNum = GetCellValue(lineWS, j, m_dataItemColumn) Then
                Do
                    ' Get lang code and check if it's given.
                    ' If not, import localized packing units.
                    ' Then localize the given text to one lang and copy it to the other.
                    ' Look if additional lines need to be localized.
                    Set dataLangCodeCell = GetCell(lineWS, j, m_dataLangCodeColumn)
                    langCode = dataLangCodeCell.Value
                    Set lineTextCell = GetCell(lineWS, j, m_lineTextColumn)
                    If LenB(langCode) = 0 Then
                        k = j
                        Do
                            k = k + 1
                            If LenB(GetCellValue(lineWS, k, m_dataLangCodeColumn)) > 0 _
                            And GetCellValue(lineWS, k, m_dataItemColumn) = itemNum Then
                                found(0) = True
                            End If
                        Loop While LenB(GetCellValue(lineWS, k, m_dataLangCodeColumn)) = 0
                        If Not found(0) Then
                            CreateNewRow lineWS, j
                            dataLangCodeCell.Value = m_translatedLangCode
                            lineTextCell.Value = GetCellValue(importWS, i, m_importTranslatedColumn)
                            CreateNewRow lineWS, j
                            dataLangCodeCell.Value = m_nativeLangCode
                            lineTextCell.Value = GetCellValue(importWS, i, m_importNativeColumn)
                            Do
                                j = j + 2
                                Set dataLangCodeCell = GetCell(lineWS, j, m_dataLangCodeColumn)
                                dataLangCodeCell.Value = m_translatedLangCode
                                AddToCellValue lineWS, j, m_lineLineNumColumn, m_baseLineNum
                                CreateNewRow lineWS, j
                                dataLangCodeCell.Value = m_nativeLangCode
                            Loop While GetCellValue(lineWS, j + 2, m_dataItemColumn) = itemNum _
                            And LenB(GetCellValue(lineWS, j + 2, m_dataLangCodeColumn)) = 0
                            found(0) = True
                            found(1) = True
                            j = j + 1
                        End If
                    ' Check which lang code is used and if correct packing unit is already input.
                    ' If not, import localized packing unit and increase all following line numbers of same locale.
                    ElseIf langCode = m_translatedLangCode Then
                        If lineTextCell.Value <> GetCellValue(importWS, i, m_importTranslatedColumn) Then
                            CreateNewRow lineWS, j
                            lineTextCell.Value = GetCellValue(importWS, i, m_importTranslatedColumn)
                            Do
                                j = j + 1
                                AddToCellValue lineWS, j, m_lineLineNumColumn, m_baseLineNum
                            Loop While GetCellValue(lineWS, j + 1, m_dataItemColumn) = itemNum _
                            And GetCellValue(lineWS, j + 1, m_dataLangCodeColumn) = langCode
                        End If
                        found(0) = True
                    ElseIf langCode = m_nativeLangCode Then
                        If lineTextCell.Value <> GetCellValue(importWS, i, m_importNativeColumn) Then
                            CreateNewRow lineWS, j
                            lineTextCell.Value = GetCellValue(importWS, i, m_importNativeColumn)
                            Do
                                j = j + 1
                                AddToCellValue lineWS, j, m_lineLineNumColumn, m_baseLineNum
                            Loop While GetCellValue(lineWS, j + 1, m_dataItemColumn) = itemNum _
                            And GetCellValue(lineWS, j + 1, m_dataLangCodeColumn) = langCode
                        End If
                        found(1) = True
                    End If
                    j = j + 1
                Loop While GetCell(lineWS, j, m_dataItemColumn).Value = itemNum
                Exit Do
            End If
            j = j + 1
        Loop
        ' If an item was not found in a specific localization before, add it at the end.
        If Not found(0) Then
            AddNewLocalizedTextLine importWS, lineWS, i, j, itemNum, m_translatedLangCode, m_importTranslatedColumn
        End If
        found(0) = False
        If Not found(1) Then
            AddNewLocalizedTextLine importWS, lineWS, i, j, itemNum, m_nativeLangCode, m_importNativeColumn
        End If
        found(1) = False
        i = i + 1
        DoEvents
    Loop
End Sub
