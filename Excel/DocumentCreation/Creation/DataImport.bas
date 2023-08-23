Attribute VB_Name = "DataImport"
Attribute VB_Description = "Handles import of values from specified file to create document."
'@Folder("DocumentCreation.Creation")
'@ModuleDescription("Handles import of values from specified file to create document.")
Option Explicit

'TODO WIP

'@EntryPoint
'@ExcelHotkey I
'@Description("Imports values from a specified file to create the document.")
Public Sub ImportDocumentData()
Attribute ImportDocumentData.VB_Description = "Imports values from a specified file to create the document."
Attribute ImportDocumentData.VB_ProcData.VB_Invoke_Func = "I\n14"
    ' Create dicts
    On Error GoTo CatchDictInitExc
    InitDicts
    On Error GoTo 0
    If False Then
CatchDictInitExc:
        ThisWorkbook.Activate
        ErrBox Err.Description
        Exit Sub
    End If
    
    ' First declarations
    Dim docWS As Worksheet
    Set docWS = ActiveWorkbook.ActiveSheet
    Dim fileID As String
    fileID = GetCellText(docWS, DocFileIDRow, DocDataColumn)
    Dim dataWBFileName As String
    dataWBFileName = Dir(DataWBsPath & fileID & FileExt)
    ' Check if file exists, exit if not.
    If IsEmpty(dataWBFileName) Then
        ErrBox FormatString(FileNotFoundError, fileID & FileExt, DataWBsPath)
        Exit Sub
    End If
    
    ' Further declarations
    Dim infoCell As Range
    Set infoCell = GetCell(docWS, DocInfoRow, DocInfosColumn)
    Dim entryIDRegex As Object
    Set entryIDRegex = CreateObject("VBScript.RegExp")
    entryIDRegex.Pattern = fileID & "\w*"
    Dim entryID As String
    Dim dataWB As Workbook
    Dim dataWS As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim currentDataEntryRow As Long
    Dim currentDocRow As Long
    Dim currentDataColumn As Long
    Dim currentDataHeaderRow As Long
    Dim currentDataHeader As String
    Dim concatDataHeader As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim categories() As String
    Dim category As Variant
    Dim categoryIDs() As String
    Dim categoryIDDoc As String
    Dim categoryIDDat As String
    Dim decimals As Long
    Dim dataCategoryFound As Boolean
    Dim translated  As Boolean
    translated = False
    Dim docCategoryValue As String
    Dim docCategoryValues() As String
    Dim dataCategoryValue As String
    Dim dataCategoryDouble As Double
    Dim categoriesNotFound As Collection
    Set categoriesNotFound = New Collection
    Dim formatStr As String
    Dim warnStr As String
    
    ' Get entry ID.
    entryID = InputBox(EntryIDQuestion, EntryIDQuestionTitle, fileID)
    If IsEmpty(entryID) Then Exit Sub
    ' Insert entry ID into document.
    SetCellValue docWS, DocEntryIDRow, DocDataColumn, entryID
    infoCell.Value = entryIDRegex.Replace(infoCell.Value, entryID)
    ' Pause for processing.
    DoEvents
    
    ' Open specified data workbook and look at first sheet.
    Set dataWB = Workbooks.Open(DataWBsPath & dataWBFileName)
    Set dataWS = dataWB.Sheets.[_Default](1)
    
    ' Get last row and column.
    lastRow = GetLastRowIndex(dataWS, DataStartingColumn)
    lastColumn = GetLastColumnIndex(dataWS, DataCategoryStartingRow)
    
    ' Start at bottom of data and work up.
    currentDataEntryRow = lastRow
    Do While currentDataEntryRow > DataCategoryStoppingRow
        ' If entry is found, process it and exit sub.
        If GetCellText(dataWS, currentDataEntryRow, DataStartingColumn) = entryID Then
            ' Go through document category rows.
            currentDocRow = DocDataStartingRow
            Do While HasBorder(GetCell(docWS, currentDocRow, DocDataColumn), xlEdgeLeft)
                ' Skip row if it hasn't got a border at the top.
                If Not HasBorder(GetCell(docWS, currentDocRow, DocCategoryColumn), xlEdgeTop) Then GoTo ContinueDo
                ' Go through categories in document row.
                docCategoryValue = vbNullString
                categories = Split(GetCellValue(docWS, currentDocRow, DocCategoryColumn), CategorySep)
                For i = 0 To UBound(categories) ' (Has to be iterator loop to not read-only-lock `categories` when exiting loop via GoTo.)
                    category = categories(i)
                    dataCategoryFound = False
                    ' Go through parts of category to find the ID.
                    ' Try 1 to at max MaxCategoryIDCombinations parts for this.
                    dataCategoryValue = vbNullString
                    categoryIDs = Split(category, IDSep)
                    For j = 0 To WorksheetFunction.Min(MaxCategoryIDCombinations - 1, UBound(categoryIDs))
                        categoryIDDoc = categoryIDs(0)
                        For k = 1 To j
                            categoryIDDoc = categoryIDDoc & IDSep & categoryIDs(k)
                        Next
                        ' Find if ID exists or should be skipped.
                        If DictDocNative.Exists(categoryIDDoc) Or DictDocTransl.Exists(categoryIDDoc) Or ListDocSkip.Contains(categoryIDDoc) Then
                            If translated Then
                                If DictDocTransl.Exists(categoryIDDat) Then
                                    categoryIDDat = DictDocTransl.Item(categoryIDDoc)
                                ElseIf DictDocNative.Exists(categoryIDDat) Then
                                    categoryIDDat = DictDocNative.Item(categoryIDDoc)
                                    translated = False
                                Else
                                    GoTo ContinueDo
                                End If
                            Else
                                If DictDocNative.Exists(categoryIDDat) Then
                                    categoryIDDat = DictDocNative.Item(categoryIDDoc)
                                ElseIf DictDocTransl.Exists(categoryIDDat) Then
                                    categoryIDDat = DictDocTransl.Item(categoryIDDoc)
                                    translated = True
                                Else
                                    GoTo ContinueDo
                                End If
                            End If
                            ' Go through data categories to find the matching one.
                            currentDataColumn = DataStartingColumn + 1
                            Do While currentDataColumn <= lastColumn
                                currentDataHeaderRow = DataCategoryStartingRow
                                concatDataHeader = vbNullString
                                Do While currentDataHeaderRow <= DataCategoryStoppingRow
                                    currentDataHeader = GetCellValue(dataWS, currentDataHeaderRow, currentDataColumn)
                                    If Not IsEmpty(currentDataHeader) Then
                                        If IsEmpty(concatDataHeader) Then
                                            concatDataHeader = currentDataHeader
                                        Else
                                            concatDataHeader = concatDataHeader & IDSep & currentDataHeader
                                        End If
                                    End If
                                    If Contains(concatDataHeader, "Example") Then
                                        Debug.Print
                                    End If
                                    If StartsWith(currentDataHeader, categoryIDDat) _
                                    Or StartsWith(concatDataHeader, categoryIDDat) Then
                                        ' Get value for category.
                                        dataCategoryValue = GetCellValue(dataWS, currentDataEntryRow, currentDataColumn)
                                        ' If it's included in the translation table, get correct translation.
                                        If DictData.Exists(dataCategoryValue) Then
                                            If translated Then
                                                dataCategoryValue = DictData.Item(dataCategoryValue)(1)
                                            Else
                                                dataCategoryValue = DictData.Item(dataCategoryValue)(0)
                                            End If
                                        ' If it's included in the decimals table, format the value accordingly as string.
                                        ElseIf DictDecimalsNative.Exists(categoryIDDoc) _
                                        Or DictDecimalsTransl.Exists(categoryIDDoc) Then
                                            If DictDecimalsNative.Exists(categoryIDDoc) Then
                                                decimals = DictDecimalsNative.Item(categoryIDDoc)
                                            Else
                                                decimals = DictDecimalsTransl.Item(categoryIDDoc)
                                            End If
                                            
                                            ' Use correct decimal symbol.
                                            If Not Contains(dataCategoryValue, DecimalSymbolNative) And Contains(dataCategoryValue, DecimalSymbolTransl) Then
                                                dataCategoryValue = Replace(dataCategoryValue, DecimalSymbolTransl, DecimalSymbolNative)
                                            End If
                                            ' Try to parse value as double.
                                            On Error GoTo CatchDoubleConvExc
                                            dataCategoryDouble = CDbl(dataCategoryValue)
                                            On Error GoTo 0
                                            ' If that works, define format.
                                            If decimals < 0 Then
                                                ' Round double value.
                                                dataCategoryDouble = Round(dataCategoryDouble, Abs(decimals))
                                                ' Set amount of decimal places.
                                                formatStr = "0." & String$(Abs(decimals), "0")
                                            Else
                                                ' Set amount of padding.
                                                formatStr = String$(decimals, "0")
                                            End If
                                            dataCategoryValue = Format$(dataCategoryDouble, formatStr)
                                            ' If not, output a warning.
                                            If False Then
CatchDoubleConvExc:
                                                WarnBox FormatString(DoubleConversionError, dataCategoryValue, category)
                                            End If
                                            ' Use appropriate decimal symbol.
                                            If translated Then
                                                dataCategoryValue = Replace(dataCategoryValue, DecimalSymbolNative, DecimalSymbolTransl)
                                            End If
                                        End If
                                        docCategoryValue = docCategoryValue & dataCategoryValue & CategorySep
                                        ' Defined category was found in data, so exit data column finder loops.
                                        dataCategoryFound = True
                                        currentDataColumn = lastColumn
                                        Exit Do
                                    End If
                                    currentDataHeaderRow = currentDataHeaderRow + 1
                                Loop
                                currentDataColumn = currentDataColumn + 1
                            Loop
                            ' Doc category was found in definitions, so exit definition finder loop.
                            Exit For
                        End If
                    Next
                    ' Check if a value was found, if not, add to "Not found" list.
                    If Not dataCategoryFound Then
                        categoriesNotFound.Add category
                    End If
                    ' Check next category.
                Next
                ' Insert correct category value into document.
                If Not IsEmpty(docCategoryValue) Then
                    docCategoryValue = RemoveLast(docCategoryValue, Len(CategorySep))
                    docCategoryValues = Split(docCategoryValue, CategorySep)
                    If UBound(docCategoryValues) > 0 Then
                        If Len(docCategoryValues(0)) > ReplaceSpaceThreshold Then
                            docCategoryValue = Replace(docCategoryValue, " ", vbNullString)
                        End If
                    End If
                End If
                SetCellValue docWS, currentDocRow, DocDataColumn, docCategoryValue
                ' Continue iteration.
ContinueDo:
                currentDocRow = currentDocRow + 1
            Loop
            ' Show warning for missing categories.
            If categoriesNotFound.Count > 0 Then
                warnStr = CategoryNotFoundWarning
                For Each category In categoriesNotFound
                    warnStr = warnStr & BulletPoint & category & vbNewLine
                Next
                WarnBox warnStr
            End If
            ' Close and exit.
            dataWB.Close
            Exit Sub
        End If
        currentDataEntryRow = currentDataEntryRow - 1
    Loop
    ' If entry isn't found, throw an error.
    ErrBox FormatString(EntryNotFoundError, entryID, dataWBFileName)
    dataWB.Close
End Sub
