Attribute VB_Name = "ExtendedTextItemDataImport"
Attribute VB_Description = "Imports extended text & boolean data for items from another Excel workbook."
'@Folder("Imports")
'@ModuleDescription("Imports extended text & boolean data for items from another Excel workbook.")
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
'@VariableDescription("Name of sheet with lines data to import into.)
Private Const m_lineSheetName As String = "ExtendedTextLine"
'@VariableDescription("Name of sheet with item data to import into.)
Private Const m_itemSheetName As String = "Item"
'@VariableDescription("Localized label of True state.")
Private Const m_trueLabel As String = "Yes"
Attribute m_trueLabel.VB_VarDescription = "Localized label of True state."
'@VariableDescription("Localized label of False state.")
Private Const m_falseLabel As String = "No"
Attribute m_falseLabel.VB_VarDescription = "Localized label of False state."
'@VariableDescription("Label for native language line.")
Private Const m_nativeLangCode As String = "EN"
Attribute m_nativeLangCode.VB_VarDescription = "Label for native language line."
'@VariableDescription("Label for translated language line.")
Private Const m_translatedLangCode As String = "DE"
Attribute m_translatedLangCode.VB_VarDescription = "Label for translated language line."

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
Private Const m_importEnglishColumn As Long = 4
Attribute m_importEnglishColumn.VB_VarDescription = "Column of native phrase in import sheet."
'@VariableDescription("Column of translated phrase in import sheet.")
Private Const m_importGermanColumn As Long = 3
Attribute m_importGermanColumn.VB_VarDescription = "Column of translated phrase in import sheet."
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
'@VariableDescription("Column of item number in item sheet.")
Private Const m_itemItemColumn As Long = 1
Attribute m_itemItemColumn.VB_VarDescription = "Column of item number in item sheet."
'@VariableDescription("Column of auto text bool in item sheet.")
Private Const m_itemAutoTextBoolColumn As Long = 2
Attribute m_itemAutoTextBoolColumn.VB_VarDescription = "Column of auto text bool in item sheet."
'@VariableDescription("The base (= 1) line number in data sheet.")
Private Const m_baseLineNum As Long = 10000
Attribute m_baseLineNum.VB_VarDescription = "The base (= 1) line number in data sheet."

' ————————————————————————————————————————————————————— '


'@Description("Inserts a new row at given position in table.")
Private Sub CreateNewRow(ByVal ws As Worksheet, ByVal row As Long)
Attribute CreateNewRow.VB_Description = "Inserts a new row at given position in table."
    ws.Rows.Item(row).Insert
    ws.Rows.Item(row + 1).Copy ws.Rows.Item(row)
End Sub

'@Description("Sets text header cells' values.")
Private Sub SetHeaderCells(ByVal ws As Worksheet, ByVal row As Long)
Attribute SetHeaderCells.VB_Description = "Sets text header cells' values."
    ws.Cells.Item(row, m_headerAllLangColumn) = m_falseLabel
    ws.Cells.Item(row, m_dataTxtNumColumn) = 0
    If ws.Cells.Item(row + 1, m_dataLangCodeColumn) = m_translatedLangCode Then
        ws.Cells.Item(row, m_dataLangCodeColumn) = m_nativeLangCode
    Else
        ws.Cells.Item(row, m_dataLangCodeColumn) = m_translatedLangCode
    End If
    If LenB(ws.Cells.Item(row, m_headerEndDateColumn)) > 0 Then
        ws.Cells.Item(row, m_headerStartDateColumn) = vbNullString
        ws.Cells.Item(row, m_headerEndDateColumn) = vbNullString
    End If
End Sub

'@Description("Adds localized line text.")
Private Sub AddNewLocalizedTextLine(ByVal importWS As Worksheet, ByVal lineWS As Worksheet, ByVal importRow As Long, ByVal lineRow As Long, ByVal itemNum As String, ByVal langCode As String, ByVal importColumn As Long)
Attribute AddNewLocalizedTextLine.VB_Description = "Adds lozalized line text."
    CreateNewRow lineWS, lineRow - 1
    lineWS.Cells.Item(lineRow, m_dataItemColumn) = itemNum
    lineWS.Cells.Item(lineRow, m_dataLangCodeColumn) = langCode
    lineWS.Cells.Item(lineRow, m_dataTxtNumColumn) = 1
    lineWS.Cells.Item(lineRow, m_lineLineNumColumn) = m_baseLineNum
    lineWS.Cells.Item(lineRow, m_lineTextColumn) = importWS.Cells.Item(importRow, importColumn)
End Sub

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Imports data for items from another excel workbook.")
Public Sub ImportItemExtendedText()
Attribute ImportItemExtendedText.VB_Description = "Imports data for items from another excel workbook."
    Dim importWB As Workbook
    Set importWB = Workbooks.Open(m_wbPath)
    Dim importWS As Worksheet
    Set importWS = importWB.Sheets.[_Default](m_importSheetName)
    Dim headerWS As Worksheet
    Set headerWS = ThisWorkbook.Sheets.[_Default](m_headerSheetName)
    Dim lineWS As Worksheet
    Set lineWS = ThisWorkbook.Sheets.[_Default](m_lineSheetName)
    Dim itemWS As Worksheet
    Set itemWS = ThisWorkbook.Sheets.[_Default](m_itemSheetName)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim itemNum As String
    Dim langCode As String
    Dim found(1) As Boolean
    
    ' Delete start date if no end date exists.
    i = m_dataStartingRow
    Do While LenB(headerWS.Cells.Item(i, m_dataItemColumn)) <> 0
        If LenB(headerWS.Cells.Item(i, m_headerStartDateColumn)) > 0 _
            And LenB(headerWS.Cells.Item(i, m_headerEndDateColumn)) = 0 Then
            headerWS.Cells.Item(i, m_headerStartDateColumn) = vbNullString
        End If
        i = i + 1
    Loop

    ' Go through import data item by item.
    i = m_importStartingRow
    Do While LenB(importWS.Cells.Item(i, m_importItemColumn)) <> 0
        itemNum = importWS.Cells.Item(i, m_importItemColumn)
        ' Find item in header data.
        j = m_dataStartingRow
        Do While LenB(headerWS.Cells.Item(j, m_dataItemColumn)) <> 0
            ' If item is found, process it and exit header loop.
            If itemNum = headerWS.Cells.Item(j, m_dataItemColumn) Then
                ' If next row isn't same item...
                If headerWS.Cells.Item(j + 1, m_dataItemColumn) <> itemNum Then
                    ' If for all langs, change to single lang and add copied row for other lang.
                    ' Else copy row and change to other lang afterwards.
                    If headerWS.Cells.Item(j, m_headerAllLangColumn) = m_trueLabel Then
                        SetHeaderCells headerWS, j
                        CreateNewRow headerWS, j
                        SetHeaderCells headerWS, j
                    Else
                        CreateNewRow headerWS, j
                        SetHeaderCells headerWS, j
                    End If
                ElseIf headerWS.Cells.Item(j, m_headerAllLangColumn) = m_trueLabel And _
                    headerWS.Cells.Item(j + 2, m_dataItemColumn) <> itemNum Then
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
            headerWS.Cells.Item(j, m_dataItemColumn) = itemNum
            CreateNewRow headerWS, j
            SetHeaderCells headerWS, j
        End If
        found(0) = False
        ' Find item in line data.
        j = m_dataStartingRow
        Do While LenB(lineWS.Cells.Item(j, m_dataItemColumn)) <> 0
            ' If item is found, process it and exit line loop.
            If itemNum = lineWS.Cells.Item(j, m_dataItemColumn) Then
                Do
                    ' Get lang code and check if it's given.
                    ' If not, localize the given text to one lang and copy it to the other.
                    ' Look if additional lines need to be localized.
                    ' Then import localized info.
                    langCode = lineWS.Cells.Item(j, m_dataLangCodeColumn)
                    If LenB(langCode) = 0 Then
                        k = j
                        Do
                            k = k + 1
                            If LenB(lineWS.Cells.Item(k, m_dataLangCodeColumn)) > 0 _
                                And lineWS.Cells.Item(k, m_dataItemColumn) = itemNum Then
                                found(0) = True
                            End If
                        Loop While LenB(lineWS.Cells.Item(k, m_dataLangCodeColumn)) = 0
                        If Not found(0) Then
                            Do
                                lineWS.Cells.Item(j, m_dataLangCodeColumn) = m_translatedLangCode
                                CreateNewRow lineWS, j
                                lineWS.Cells.Item(j, m_dataLangCodeColumn) = m_nativeLangCode
                                j = j + 2
                            Loop While lineWS.Cells.Item(j, m_dataItemColumn) = itemNum _
                                And LenB(lineWS.Cells.Item(j, m_dataLangCodeColumn)) = 0
                            CreateNewRow lineWS, j - 1
                            lineWS.Cells.Item(j, m_dataLangCodeColumn) = m_translatedLangCode
                            lineWS.Cells.Item(j, m_lineTextColumn) = importWS.Cells.Item(i, m_importEnglishColumn)
                            lineWS.Cells.Item(j, m_lineLineNumColumn) = lineWS.Cells.Item(j, m_lineLineNumColumn) + m_baseLineNum
                            CreateNewRow lineWS, j
                            lineWS.Cells.Item(j, m_dataLangCodeColumn) = m_nativeLangCode
                            lineWS.Cells.Item(j, m_lineTextColumn) = importWS.Cells.Item(i, m_importGermanColumn)
                            found(0) = True
                            found(1) = True
                            j = j + 1
                        End If
                    ' Check which lang code is used and if correct info is already input.
                    ' If not, import localized info.
                    ElseIf langCode = m_translatedLangCode Then
                        Do While lineWS.Cells.Item(j + 1, m_dataItemColumn) = itemNum _
                            And lineWS.Cells.Item(j + 1, m_dataLangCodeColumn) = langCode
                            j = j + 1
                        Loop
                        If lineWS.Cells.Item(j, m_lineTextColumn) <> importWS.Cells.Item(i, m_importEnglishColumn) Then
                            CreateNewRow lineWS, j
                            j = j + 1
                            lineWS.Cells.Item(j, m_lineTextColumn) = importWS.Cells.Item(i, m_importEnglishColumn)
                            lineWS.Cells.Item(j, m_lineLineNumColumn) = lineWS.Cells.Item(j, m_lineLineNumColumn) + m_baseLineNum
                        End If
                        found(0) = True
                    ElseIf langCode = m_nativeLangCode Then
                        Do While lineWS.Cells.Item(j + 1, m_dataItemColumn) = itemNum _
                            And lineWS.Cells.Item(j + 1, m_dataLangCodeColumn) = langCode
                            j = j + 1
                        Loop
                        If lineWS.Cells.Item(j, m_lineTextColumn) <> importWS.Cells.Item(i, m_importGermanColumn) Then
                            CreateNewRow lineWS, j
                            j = j + 1
                            lineWS.Cells.Item(j, m_lineTextColumn) = importWS.Cells.Item(i, m_importGermanColumn)
                            lineWS.Cells.Item(j, m_lineLineNumColumn) = lineWS.Cells.Item(j, m_lineLineNumColumn) + m_baseLineNum
                        End If
                        found(1) = True
                    End If
                    j = j + 1
                Loop While lineWS.Cells.Item(j, m_dataItemColumn) = itemNum
                Exit Do
            End If
            j = j + 1
        Loop
        ' If an item was not found in a specific localization before, add it at the end.
        If Not found(0) Then
            AddNewLocalizedTextLine importWS, lineWS, i, j, itemNum, m_translatedLangCode, m_importEnglishColumn
        End If
        found(0) = False
        If Not found(1) Then
            AddNewLocalizedTextLine importWS, lineWS, i, j, itemNum, m_nativeLangCode, m_importGermanColumn
        End If
        found(1) = False
        ' Find item in item data.
        j = m_dataStartingRow
        Do While LenB(itemWS.Cells.Item(j, m_itemItemColumn)) <> 0
            ' If item is found, process it and exit item loop.
            If itemNum = itemWS.Cells.Item(j, m_itemItemColumn) Then
                itemWS.Cells.Item(j, m_itemAutoTextBoolColumn) = m_trueLabel
                found(0) = True
                Exit Do
            End If
            j = j + 1
        Loop
        If Not found(0) Then
            CreateNewRow itemWS, j - 1
            itemWS.Cells.Item(j, m_itemItemColumn) = itemNum
            itemWS.Cells.Item(j, m_itemAutoTextBoolColumn) = m_trueLabel
        End If
        found(0) = False
        i = i + 1
        DoEvents
    Loop
End Sub
