Attribute VB_Name = "CategoryContinuation"
Attribute VB_Description = "Handles category names repeated at later part of file."
'@Folder("DocumentCreation.Unification")
'@ModuleDescription("Handles category names repeated at later part of file.")
Option Explicit

' Runtime constants
' Change each run to get results you want.
'@VariableDescription("Wether to adjust formatting of header and category cells, too. What that formatting is has to be defined in code.")
Private Const m_adjustFormat As Boolean = True
Attribute m_adjustFormat.VB_VarDescription = "Wether to adjust formatting of header and category cells, too. What that formatting is has to be defined in code."
'@VariableDescription("Wether to also replace consecutive headers with first header.")
Private Const m_replaceHeaders As Boolean = False
Attribute m_replaceHeaders.VB_VarDescription = "Wether to also replace consecutive headers with first header."
'@VariableDescription("Reverse the continuation order to use the topmost header as a basis and fill the other ones accordingly.")
Private Const m_reverseOrder As Boolean = False
Attribute m_reverseOrder.VB_VarDescription = "Reverse the continuation order to use the topmost header as a basis and fill the other ones accordingly."

' ————————————————————————————————————————————————————— '

'@Description("Adjusts format of header cells.")
Private Function AdjustHeaderFormat(ByVal cell_ As Range, ByVal changed As Boolean) As Boolean
Attribute AdjustHeaderFormat.VB_Description = "Adjusts format of header cells."
    Dim changed_ As Boolean
    changed_ = changed
    If Not cell_.Font.Bold Then
        cell_.Font.Bold = True
        changed_ = True
    End If
    AdjustHeaderFormat = changed_
End Function

'@Description("Adjusts format of category cells.")
Private Function AdjustCategoryFormat(ByVal cell_ As Range, ByVal changed As Boolean) As Boolean
Attribute AdjustCategoryFormat.VB_Description = "Adjusts format of category cells."
    Dim changed_ As Boolean
    changed_ = changed
    If Not IsEmpty(cell_.Value) Then
        If Not cell_.Font.Bold Then
            cell_.Font.Bold = True
            changed_ = True
        End If
        If cell_.Column > 2 And cell_.HorizontalAlignment <> xlCenter Then
            cell_.HorizontalAlignment = xlCenter
            changed_ = True
        End If
    End If
    AdjustCategoryFormat = changed_
End Function

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Equalizes category names repeated at a later part of the files to ones at top.")
Public Sub EnsureContinuedCategoriesCorrectHeaderAndFormatting()
Attribute EnsureContinuedCategoriesCorrectHeaderAndFormatting.VB_Description = "Equalizes category names repeated at a later part of the files to ones at top."
    ' Declarations
    Dim dataWBFileName As String
    dataWBFileName = Dir(DataWBsPath & AllDataWBFilesPattern)
    Dim ignore As Variant
    Dim ignores() As String
    ignores = Split(IgnoreList, ListSep)
    Dim dataWB As Workbook
    Dim dataWS As Worksheet
    Dim changed As Boolean
    changed = False
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim i As Long
    Dim currentRow As Long
    Dim currentColumn As Long
    Dim currentCell As Range
    Dim header As String
    Dim headerMarker As String
    Dim headerRows As Collection
    Dim headerRow As Variant
    Dim headerCategoryGap As Long
    headerCategoryGap = DataCategoryStartingRow - DataHeaderRow
    Dim dataLastCategoryStartingRow As Long
    Dim categoryPart As String
    
    ' Go through all files fitting the pattern.
    Do Until IsEmpty(dataWBFileName)
        ' If file name contains ignore marker symbol, skip it, continuing with next file.
        For Each ignore In ignores
            If Contains(dataWBFileName, CStr(ignore)) _
            Then GoTo Skip
        Next
        ' Open current workbook and look at first sheet.
        Set dataWB = Workbooks.Open(DataWBsPath & dataWBFileName)
        Set dataWS = dataWB.Sheets.[_Default](1)
        ' If first sheet contains the first ignore symbol, move second sheet before first and look at that sheet.
        changed = ChooseCorrectSheet(dataWB, dataWS, ignores(0))
        ' Get last row and column.
        lastRow = GetLastRowIndex(dataWS, DataStartingColumn)
        lastColumn = GetLastColumnIndex(dataWS, DataCategoryStartingRow)
        ' Get header to search.
        i = 0
        currentRow = DataHeaderRow - 1
        Do
            If i Mod lastColumn = 0 Then
                currentRow = currentRow + 1
            End If
            currentColumn = DataHeaderColumn + (i Mod lastColumn)
            header = GetCellValue(dataWS, currentRow, currentColumn)
            i = i + 1
        Loop While IsEmpty(header) And currentRow < DataCategoryStartingRow
        ' If header is not in correct place, move it there.
        If currentRow = DataCategoryStartingRow Then
            ErrBox MissingHeaderError
            Exit Sub
        ElseIf currentColumn > DataHeaderColumn Or currentRow > DataHeaderRow Then
            GetCell(dataWS, DataHeaderRow, DataHeaderColumn).Value = header
            GetCell(dataWS, currentRow, currentColumn).Value = vbNullString
            changed = True
        End If
        ' Format first header.
        If m_adjustFormat Then
            Set currentCell = GetCell(dataWS, DataHeaderRow, DataHeaderColumn)
            changed = AdjustHeaderFormat(currentCell, changed)
        End If
        ' Get header marker.
        headerMarker = Split(header, IDSep)(0) & IDSep
        ' Reset header rows collection.
        Set headerRows = New Collection
        ' Find all rows with header after the first one and put their indices in a collection.
        currentRow = DataCategoryStoppingRow + 1
        Do While currentRow <= lastRow
            If StartsWith(GetCellValue(dataWS, currentRow, currentColumn), headerMarker) Then
                headerRows.Add currentRow
                ' Move consecutive headers if in wrong column.
                If currentColumn > DataHeaderColumn Then
                    GetCell(dataWS, currentRow, DataHeaderColumn).Value = GetCellValue(dataWS, currentRow, currentColumn)
                    GetCell(dataWS, currentRow, currentColumn).Value = vbNullString
                End If
                ' Format consecutive headers.
                If m_adjustFormat Then
                    Set currentCell = GetCell(dataWS, currentRow, DataHeaderColumn)
                    changed = AdjustHeaderFormat(currentCell, changed)
                End If
            End If
            currentRow = currentRow + 1
        Loop
        ' If there is more than one header in file or formatting needs adjustment, continue and/or format categories appropriately.
        If headerRows.Count > 0 Or m_adjustFormat Then
            currentColumn = DataStartingColumn
            If m_reverseOrder Then
                If headerRows.Count > 0 Then
                    header = GetCellValue(dataWS, headerRows.Item(headerRows.Count), DataHeaderColumn)
                    dataLastCategoryStartingRow = headerRows.Item(headerRows.Count) + headerCategoryGap
                    headerRows.Remove headerRows.Count
                    headerRows.Add DataHeaderRow
                Else
                    header = GetCellValue(dataWS, DataHeaderRow, DataHeaderColumn)
                    dataLastCategoryStartingRow = DataCategoryStartingRow
                End If
            Else
                header = GetCellValue(dataWS, DataHeaderRow, DataHeaderColumn)
            End If
            ' Change all categories after all headers to the first one's.
            Do While currentColumn <= lastColumn
                i = 0
                Do
                    ' Get first categories.
                    If m_reverseOrder Then
                        currentRow = dataLastCategoryStartingRow + i
                    Else
                        currentRow = DataCategoryStartingRow + i
                    End If
                    categoryPart = GetCellValue(dataWS, currentRow, currentColumn)
                    ' Format first categories.
                    If m_adjustFormat Then
                        Set currentCell = GetCell(dataWS, currentRow, currentColumn)
                        changed = AdjustCategoryFormat(currentCell, changed)
                    End If
                    ' Iterate all category appearances.
                    For Each headerRow In headerRows
                        ' Set consecutive headers.
                        If m_replaceHeaders Then
                            If GetCellValue(dataWS, headerRow, DataHeaderColumn) <> header Then
                                GetCell(dataWS, headerRow, DataHeaderColumn).Value = header
                                changed = True
                            End If
                        End If
                        ' Set consecutive categories.
                        currentRow = headerRow + headerCategoryGap + i
                        If GetCellValue(dataWS, currentRow, currentColumn) <> categoryPart Then
                            GetCell(dataWS, currentRow, currentColumn).Value = categoryPart
                            changed = True
                        End If
                        ' Format consecutive categories.
                        If m_adjustFormat Then
                            Set currentCell = GetCell(dataWS, currentRow, currentColumn)
                            changed = AdjustCategoryFormat(currentCell, changed)
                        End If
                    Next
                    i = i + 1
                Loop While i <= DataCategoryStoppingRow - DataCategoryStartingRow
                currentColumn = currentColumn + 1
            Loop
        End If
        ' Finish processing.
        If changed Then
            dataWB.Save
            changed = False
        End If
        dataWB.Close
Skip:
        dataWBFileName = Dir
        DoEvents
    Loop
End Sub
