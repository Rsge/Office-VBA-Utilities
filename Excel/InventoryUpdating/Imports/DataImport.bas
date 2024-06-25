Attribute VB_Name = "DataImport"
Attribute VB_Description = "Import weight data from given CSV files."
'@Folder("InventoryUpdating.Imports")
'@ModuleDescription("Import weight data from given CSV files.")
Option Explicit

'@Description("Gets path of data import files or export file and determines if file(s) is/are available.")
Private Function GetFilePath(ByVal pathCell As Range, ByVal forImport As Boolean) As String
Attribute GetFilePath.VB_Description = "Gets path of data import files or export file and determines if file(s) is/are available."
    ' Variables
    Dim filePath As String
    filePath = pathCell.Value
    Dim repeated As Boolean
    Dim label As String
    If forImport Then
        label = ImportLabel
    Else ' If for export...
        label = ExportLabel
        ' Get export file name.
        Dim exportFileName As String
        exportFileName = BuildWBName(ExportDateFormat)
    End If
    ' Loop as long as no valid path is found and not canceled.
    Dim FolderDialog As FileDialog
    Do
        ' If no path specified, define one.
        If IsEmpty(filePath) Then
            ' MsgBox to cancel folder dialog
            If Not repeated Then
                If MsgBoxCanceled(FormatString(NoPathWarning, label)) Then Exit Function
            End If
            ' Open folder dialog.
            Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
            If FolderDialog.Show = 0 Then Exit Function
            ' Get path.
            filePath = FolderDialog.SelectedItems.Item(1) & Application.PathSeparator
            pathCell.Value = filePath
        End If
        ' Check file existence.
        If (forImport And IsEmpty(Dir(filePath & DataFilePattern))) _
        Or (Not forImport And IsEmpty(Dir(filePath & exportFileName))) Then
            ' MsgBox to cancel repeated folder dialog
            If MsgBoxCanceled(FormatString(NoFilesWarning, label)) Then Exit Function
            repeated = True
            filePath = vbNullString
        Else
            Exit Do
        End If
    Loop
    GetFilePath = filePath
End Function

'@EntryPoint
'@Description("Imports weighing data from given data files.")
Public Sub ImportDataFiles()
Attribute ImportDataFiles.VB_Description = "Imports weighing data from given data files."
    ' Define paths, names and the missing items list.
    Dim missingItems As Object
    Set missingItems = CreateObject("System.Collections.ArrayList")
    Dim importDataFilesPath As String
    importDataFilesPath = GetFilePath(GetActCell(ImportPathAndResetMarkerRow, PathCellsColumn), forImport:=True)
    If IsEmpty(importDataFilesPath) Then Exit Sub
    Dim exportWBPath As String
    exportWBPath = GetFilePath(GetActCell(ExportPathRow, PathCellsColumn), forImport:=False)
    If IsEmpty(exportWBPath) Then Exit Sub
    Dim exportWBName As String
    exportWBName = BuildWBName(ExportDateFormat)
    Dim exportWB As Workbook
    ' Select first sheet.
    ActiveWorkbook.Worksheets.Item(1).Select
    ' Create copy of this workbook with current date and use that from now on, if not already in use.
    Dim currentWB As Workbook
    Dim isNew As Boolean
    If CreateWBCopy Then
        Dim newName As String
        newName = BuildWBName(ActFileDateFormat, isMakroWB:=True)
        isNew = ThisWorkbook.Name <> newName
        Dim newFilePath As String
        newFilePath = ThisWorkbook.Path & Application.PathSeparator & newName
        If isNew Then
            ThisWorkbook.SaveCopyAs newFilePath
            Set currentWB = Workbooks.Open(newFilePath)
        Else
            Set currentWB = ThisWorkbook
        End If
    Else
        isNew = False
        Set currentWB = ThisWorkbook
    End If
    ' Open export workbook, make sure it's not read-only and put it in background.
    Set exportWB = Workbooks.Open(exportWBPath & exportWBName)
    If exportWB.ReadOnly Then
        WarnBox ReadOnlyWarning
        exportWB.Close
        Exit Sub
    End If
    Dim exportSheet As Worksheet
    Set exportSheet = exportWB.ActiveSheet
    currentWB.Activate
    ' Import baseline from export file, if not done already.
    If IsEmpty(GetActCellValue(ImportPathAndResetMarkerRow, ResetMarkerColumn)) Then
        exportSheet.Range(DataRegionStartCell).CurrentRegion.Copy ActiveSheet.Range(DataRegionStartCell)
    Else
        SetActCellValue ImportPathAndResetMarkerRow, ResetMarkerColumn, vbNullString
    End If
    ' Back up worksheet.
    Dim sheetName As String
    sheetName = BackupSheetLabel & Format$(Now, DataDateFormat)
    If Not ContainsSheetStartingWith(sheetName) Then
        ' Delete old backups.
        DeleteSheetsStartingWith BackupSheetLabel
        ' Create new backup.
        ActiveSheet.Copy After:=ActiveSheet
        ActiveSheet.Name = sheetName
    Else
        WarnBox DoneAlreadyWarning
        exportWB.Close
        Exit Sub
    End If
    ActiveWorkbook.Sheets.Item(1).Select
    ' Create blacklist and special item list.
    Dim blacklist As Object
    Dim specialItemsList As Object
    With ActiveWorkbook.Sheets.Item(DefSheetName)
        Set blacklist = GetTableAsList(.Range(BlacklistedItemsTableName))
        Set specialItemsList = GetTableAsList(.Range(SpecialItemsTableName))
    End With
    ' Iterate over all items' files in data file folder.
    Dim file As Object
    Dim itemNum As String
    Dim hasDuplicate As Boolean
    Dim isSpecialItem As Boolean
    Dim itemColumnRange As Range
    Dim itemCell As Range
    Dim firstCellAdr As String
    Dim missesInTable As Boolean
    Dim doRepeat As Boolean
    Dim itemRow As Long
    Dim description As String
    Dim descHasMarker As Boolean
    Dim actValueRange As Range
    Dim i As Long
    Dim importData() As String
    Dim currentAmount As Double
    Dim previousAmount As Double
    Dim diff As Double
    Dim unit As String
    Dim importBBDateStr As String
    For Each file In CreateObject("Scripting.FileSystemObject").GetFolder(importDataFilesPath).Files
        ' Get item.
        itemNum = GetFileNameWithoutExtension(file)
        ' Account for special, duplicate items.
        hasDuplicate = False
        isSpecialItem = False
        If specialItemsList.Contains(Replace(itemNum, SpecialItemFileMarker, vbNullString)) Then
            hasDuplicate = True
            If EndsWith(itemNum, SpecialItemFileMarker) Then
                isSpecialItem = True
                itemNum = Replace(itemNum, SpecialItemFileMarker, vbNullString)
            End If
        End If
        ' Don't process blacklisted items.
        If blacklist.Contains(itemNum) Then GoTo Continue
        ' Find item's cell.
        Set itemColumnRange = ActiveSheet.Columns.Item(ItemColumn)
        Set itemCell = itemColumnRange.Find(itemNum)
        ' Find item's cell. If not found, add to missing and create it.
        If Not itemCell Is Nothing Then
            missesInTable = False
            firstCellAdr = itemCell.Address
            ' Account for special items having two entries.
            Do
                itemRow = itemCell.Row
                doRepeat = False
                ' If item has a special duplicate, the description differentiates the entries.
                If hasDuplicate Then
                    description = GetActCellValue(itemRow, DescriptionColumn)
                    descHasMarker = StartsWith(description, SpecialItemDescriptionMarker)
                    ' If the description isn't for the current item's variant, try the next result.
                    ' Otherwise go on with the import.
                    If isSpecialItem Xor descHasMarker Then
                        Set itemCell = itemColumnRange.FindNext(itemCell)
                        If itemCell.Address = firstCellAdr Then
                            missesInTable = True
                            Exit Do
                        End If
                        doRepeat = True
                    End If
                End If
            Loop While doRepeat
        Else
            missesInTable = True
        End If
        If missesInTable Then
            ' Add row for missing item.
            missingItems.Add itemNum
            itemRow = StartingRow
            Do Until IsEmpty(GetActCellValue(itemRow, ItemColumn))
                ' Find where the item belongs.
                If StrComp(itemNum, GetActCellValue(itemRow, ItemColumn)) = -1 Then
                    Exit Do
                End If
                itemRow = itemRow + 1
            Loop
            If itemRow > StartingRow Then
                CreateNewActRow itemRow, copyFrom:=-1
            Else
                CreateNewActRow itemRow, copyFrom:=1
            End If
            ' Get missing item's data or set it to default values.
            importData = Split(GetFirstLine(file.Path, 2)(1), Sep)
            SetActCellValue itemRow, ItemColumn, itemNum
            If isSpecialItem Then
                SetActCellValue itemRow, DescriptionColumn, SpecialItemDescriptionMarker & Space$(1)
            Else
                SetActCellValue itemRow, DescriptionColumn, vbNullString
            End If
            importBBDateStr = importData(ImportsCurrentBBDateColumn)
            SetActCellValue itemRow, BBDateColumn, importBBDateStr
            GetActCell(itemRow, BBDateColumn).NumberFormat = DataDateFormat
            currentAmount = CDbl(Replace(importData(ImportsCurrentAmountColumn), ImportUnit, vbNullString))
            unit = Replace(ImportUnit, Space$(1), vbNullString)
            If currentAmount >= UnitSwitchAmount Then
                SetActCellValue itemRow, UnitColumn, KiloUnitPrefix & unit
                SetActCellValue itemRow, PreviousAmountColum, currentAmount / 1000
            Else
                SetActCellValue itemRow, UnitColumn, unit
                SetActCellValue itemRow, PreviousAmountColum, currentAmount
            End If
            SetActCellValue itemRow, AmountDiffColumn, 0
            SetActCellValue itemRow, LastChangedDateColumn, PlaceholderDate
            Set actValueRange = ActiveSheet.Range(DataRegionStartCell).CurrentRegion
            For i = actValueRange.Columns.Count To ActiveSheet.Range(DataRegionStartCell).Column + LastChangedDateColumn Step -1
                actValueRange.Cells.Item(itemRow - ActiveSheet.Range(DataRegionStartCell).Row + 1, i).Value = vbNullString
            Next
        End If
        ' Process item's data.
        importData = Split(GetLastLine(file.Path)(0), Sep)
        If UBound(importData) < ImportsLastDataColumn - 1 Then
            ' If data has wrong format, show error and reset workbook.
            MsgBox FormatString(FormattingError, file.Name)
            exportWB.Close SaveChanges:=False
            ResetTable
            ' Close original workbook if copy was created.
            If CreateWBCopy Then If isNew Then ThisWorkbook.Close
            Exit Sub
        End If
        ' Account for kilo-unit.
        currentAmount = Replace(importData(ImportsCurrentAmountColumn), ImportUnit, vbNullString)
        unit = GetActCellValue(itemRow, UnitColumn)
        If Contains(unit, KiloUnitPrefix) Or unit = LitersUnit Then
            currentAmount = currentAmount / 1000
        End If
        ' Change data in Excel table only if imported data is newer.
        If CDate(GetActCellValue(itemRow, LastChangedDateColumn)) < CDate(importData(ImportsLastChangedDateColumn)) Then
            ' BB date
            importBBDateStr = importData(ImportsCurrentBBDateColumn)
            If importBBDateStr = PlaceholderDate Then
                SetActCellValue itemRow, BBDateColumn, vbNullString
            Else
                On Error Resume Next
                SetActCellValue itemRow, BBDateColumn, CDate(importBBDateStr)
                On Error GoTo 0
            End If
            GetActCell(itemRow, BBDateColumn).NumberFormat = DataDateFormat
            ' Last changed date
            SetActCellValue itemRow, LastChangedDateColumn, Date
            GetActCell(itemRow, LastChangedDateColumn).NumberFormat = DataDateFormat
            ' Amount
            previousAmount = GetActCellValue(itemRow, NewAmountColumn)
            SetActCellValue itemRow, PreviousAmountColum, previousAmount
            diff = Math.Round(currentAmount - previousAmount, Decimals)
            SetActCellValue itemRow, AmountDiffColumn, diff
        End If
Continue:
    Next
    ' Export to export file.
    With ActiveSheet.Range(DataRegionStartCell)
        .CurrentRegion.Copy exportSheet.Range(DataRegionStartCell)
        .Select
    End With
    ' Save.
    exportWB.Save
    currentWB.Save
    MsgBox SuccessInfo
    ' Show missing items list or just close export workbook.
    If missingItems.Count > 0 Then
        Dim missingItemNum As Variant
        Dim missingItemsListString As String
        missingItemsListString = EntryNotAvailableWarning
        For Each missingItemNum In missingItems
            missingItemsListString = missingItemsListString & missingItemNum & vbNewLine
        Next
        MsgBox missingItemsListString
    Else
        exportWB.Close
    End If
    ' Close original workbook if copy was created.
    If CreateWBCopy Then If isNew Then ThisWorkbook.Close
End Sub
