Attribute VB_Name = "ViewUnification"
Attribute VB_Description = "Handles unification of view settings in data files."
'@Folder("DocumentCreation.Unification")
'@ModuleDescription("Handles unification of view settings in data files.")
Option Explicit

'@Description("Defines what view settings to use.")
Private Function UnifyView(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastHeaderRow As Long, ByVal gap As Long) As Boolean
Attribute UnifyView.VB_Description = "Defines what view settings to use."
    Dim firstDataRow As Long
    firstDataRow = lastHeaderRow + gap + 1
    ws.Activate
    With ActiveWindow
        If .Split Then
            .FreezePanes = False
            .Split = False
        End If
        UnifyView = RemoveSpaceOnly(ws, firstDataRow - 1, DataStartingColumn)
        If Not IsEmpty(GetCellValue(ws, firstDataRow - 1, DataStartingColumn)) Then
            firstDataRow = firstDataRow - 1
        End If
        GetCell(ws, firstDataRow, DataStartingColumn).Select
        .ScrollColumn = DataStartingColumn
        .ScrollRow = lastHeaderRow
        .FreezePanes = True
        GetCell(ws, lastRow, DataStartingColumn).Select
        .ScrollRow = lastRow - ShowRowsCount + 1
    End With
    UnifyView = True
End Function

'@EntryPoint
'@Description("Unifies view settings for data files.")
Public Sub UnifyViewSettings()
Attribute UnifyViewSettings.VB_Description = "Unifies view settings for data files."
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
    Dim header As String
    Dim lastRow As Long
    Dim lastHeaderRow As Long
    Dim currentRow As Long
    Dim currentCell As Range
    Dim headerCategoryGap As Long
    headerCategoryGap = DataCategoryStartingRow - DataHeaderRow
    Dim categoryLength As Long
    categoryLength = DataCategoryStoppingRow - DataCategoryStartingRow
    Dim freezeGap As Long
    freezeGap = headerCategoryGap + categoryLength + 1
    
    'Go through all files fitting the pattern.
    Do Until IsEmpty(dataWBFileName)
        ' If file name contains ignore marker symbol, skip it, continuing with next file.
        For Each ignore In ignores
            If Contains(dataWBFileName, CStr(ignore)) _
            Then GoTo Skip
        Next
        ' Open current workbook and look at first sheet.
        Set dataWB = Workbooks.Open(DataWBsPath & dataWBFileName)
        Set dataWS = dataWB.Sheets.[_Default](1)
        ' Get header.
        header = GetCellValue(dataWS, DataHeaderRow, DataHeaderColumn)
        ' Get last used row.
        lastRow = GetLastRowIndex(dataWS, DataStartingColumn)
        currentRow = lastRow
        Set currentCell = GetCell(dataWS, currentRow, DataStartingColumn)
        Do Until currentRow = DataCategoryStoppingRow
            ' If no spaces were removed and the cell is not empty, exit loop.
            If RemoveSpaceOnly(dataWS, currentRow, DataStartingColumn) Then
                changed = True
            ElseIf Not IsEmpty(currentCell.Value) Then
                Exit Do
            End If
            currentRow = currentRow - 1
            Set currentCell = GetCell(dataWS, currentRow, DataStartingColumn)
        Loop
        lastRow = currentRow
        ' Get last header row.
        lastHeaderRow = DataHeaderRow
        Do Until currentRow = DataCategoryStoppingRow
            If currentCell.Value = header Then
                lastHeaderRow = currentRow
                Exit Do
            End If
            currentRow = currentRow - 1
            Set currentCell = GetCell(dataWS, currentRow, DataStartingColumn)
        Loop
        ' Do view changes.
        If UnifyView(dataWS, lastRow, lastHeaderRow, freezeGap) Then
            changed = True
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
