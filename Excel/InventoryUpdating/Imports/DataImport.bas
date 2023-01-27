Attribute VB_Name = "DataImport"
Attribute VB_Description = "Imports weight data from given CSV files."
'@Folder("InventoryUpdating.Imports")
'@ModuleDescription("Imports weight data from given CSV files.")
Option Explicit

'@VariableDescription "Warning for a file's item number not being present in table."
Private Const m_entryNotAvailableWarning As String = "No entry exists for the following items:" & vbNewLine
Attribute m_entryNotAvailableWarning.VB_VarDescription = "Warning for a file's item number not being present in table."
'@VariableDescription "Info about successful import."
Private Const m_successInfo As String = "Data import completed successfully."
Attribute m_successInfo.VB_VarDescription = "Info about successful import."
'@VariableDescription "Warning about import already done."
Private Const m_doneAlreadyWarning As String = "Data import was already carried out today."
Attribute m_doneAlreadyWarning.VB_VarDescription = "Warning about import already done."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Imports weighing data from given data files.")
Public Sub ImportDataFiles()
Attribute ImportDataFiles.VB_Description = "Imports weighing data from given data files."
    ' Variables
    Dim dataFilePath As String
    dataFilePath = GetDataFilePath(ActiveSheet.Cells(PathCellRow, PathCellColumn))
    If IsEmpty(dataFilePath) Then Exit Sub
    Dim missingItems As Object
    Set missingItems = CreateObject("System.Collections.ArrayList")
    ' Backup worksheet.
    ActiveSheet.Copy After:=ActiveSheet
    On Error GoTo ErrorHandler
    ActiveSheet.Name = BackupLabel & Format$(Now, DateFormat)
    On Error GoTo 0
    ActiveWorkbook.Sheets.Item(1).Select
    ' Iterate over all items' files in data file folder.
    Dim file As Object
    For Each file In CreateObject("Scripting.FileSystemObject").GetFolder(dataFilePath).Files
        ' Gett item.
        Dim itemNum As String
        itemNum = GetFileNameWithoutExtension(file)
        ' Account for special, duplicate items.
        Dim hasDuplicate As Boolean
        hasDuplicate = False
        Dim isSpecialItem As Boolean
        isSpecialItem = False
        If Contains(SpecialItems, itemNum) Or EndsWith(itemNum, SpecialItemFileMarker) Then
            hasDuplicate = True
            If EndsWith(itemNum, SpecialItemFileMarker) Then
                isSpecialItem = True
                itemNum = Replace(itemNum, SpecialItemFileMarker, vbNullString)
            End If
        End If
        ' Find item's cell.
        Dim itemColumnRange As Range
        Set itemColumnRange = ActiveSheet.Columns(ItemColumn)
        Dim itemCell As Range
        Set itemCell = itemColumnRange.Find(itemNum)
        ' Process item's data if cell is found, otherwise add it to missing list.
        If Not itemCell Is Nothing Then
Retry:
            Dim itemRow As Long
            itemRow = itemCell.Row
            If hasDuplicate Then
                Dim description As String
                description = ActiveSheet.Cells(itemRow, DescriptionColumn).Value
                Dim descHasMarker As Boolean
                descHasMarker = StartsWith(description, SpecialItemDescriptionMarker)
                If Not ((isSpecialItem And descHasMarker) Or (Not isSpecialItem And Not descHasMarker)) Then
                    Set itemCell = itemColumnRange.FindNext(itemCell)
                    If itemCell Is Nothing Then GoTo MissingItem
                    ' HACK It's easiest like this - probably could do it "better", but it works...
                    GoTo Retry
                End If
            End If
            Dim ImportData() As String
            ImportData = Split(GetLastLine(file.path)(0), Sep)
            ' Account for kilo-unit.
            Dim currentAmount As Double
            currentAmount = Replace(ImportData(ImportsCurrentAmountColumn), ImportUnit, vbNullString)
            Dim unit As String
            unit = ActiveSheet.Cells(itemRow, UnitColumn).Value
            If Contains(unit, KiloUnitPrefix) Or unit = LitersUnit Then
                currentAmount = currentAmount / 1000
            End If
            ' Change data in Excel table only if imported data is newer.
            If CDate(ActiveSheet.Cells(itemRow, LastChangedDateColumn).Value) < CDate(ImportData(ImportsLastChangedDateColumn)) Then
                ' BB date
                Dim currentBBDateStr As String
                currentBBDateStr = ImportData(ImportsCurrentBBDateColumn)
                If currentBBDateStr = PlaceholderDate Then
                    ActiveSheet.Cells(itemRow, BBDateColumn).Value = vbNullString
                Else
                    On Error Resume Next
                    ActiveSheet.Cells(itemRow, BBDateColumn).Value = CDate(currentBBDateStr)
                    On Error GoTo 0
                End If
                ' Last changed date
                ActiveSheet.Cells(itemRow, LastChangedDateColumn).Value = Now
                ' Amount
                Dim previousAmount As Double
                previousAmount = ActiveSheet.Cells(itemRow, NewAmountColumn).Value
                ActiveSheet.Cells(itemRow, PreviousAmountColum).Value = previousAmount
                Dim diff As Double
                diff = Math.Round(currentAmount - previousAmount, Decimals)
                ActiveSheet.Cells(itemRow, AmountDiffColumn).Value = diff
            End If
MissingItem:
        ElseIf Not Contains(BlacklistedItems, itemNum) Then
                missingItems.Add itemNum
        End If
    Next
    MsgBox (m_successInfo)
    ' Show missing items list.
    If missingItems.Count > 0 Then
        Dim missingItemNum As Variant
        Dim missingItemsListString As String
        missingItemsListString = m_entryNotAvailableWarning
        For Each missingItemNum In missingItems
            missingItemsListString = missingItemsListString & missingItemNum & vbNewLine
        Next
        MsgBox (missingItemsListString)
    End If
    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then
        Err.Clear
        Application.DisplayAlerts = False
        ActiveWorkbook.ActiveSheet.Delete
        Application.DisplayAlerts = True
        ActiveWorkbook.Sheets.Item(1).Select
        MsgBox m_doneAlreadyWarning, vbExclamation
    End If
End Sub
