Attribute VB_Name = "DataImporting"
Attribute VB_Description = "Imports weight data from given CSV files."
'@Folder "Inventory updating"
'@ModuleDescription "Imports weight data from given CSV files."
Option Explicit

'@VariableDescription "Warning for a file's item number not being present in table."
Private Const EntryNotAvailableWarning As String = "No entry exists for the following items:" & vbNewLine
Attribute EntryNotAvailableWarning.VB_VarDescription = "Warning for a file's item number not being present in table."
'@VariableDescription "Info about successful import."
Private Const SuccessInfo As String = "Data import completed successfully."
Attribute SuccessInfo.VB_VarDescription = "Info about successful import."
'@VariableDescription "Warning about import already done."
Private Const DoneAlreadyWarning As String = "Data import was already carried out today."
Attribute DoneAlreadyWarning.VB_VarDescription = "Warning about import already done."


'@EntryPoint
'@Description "Imports weighing data from given data files."
Public Sub ImportDataFiles()
Attribute ImportDataFiles.VB_Description = "Imports weighing data from given data files."
    'Variables
    Dim DataFilePath As String
    DataFilePath = GetDataFilePath(ActiveSheet.Cells(PathCellRow, PathCellColumn))
    If LenB(DataFilePath) = 0 Then Exit Sub
    Dim MissingItems As Object
    Set MissingItems = CreateObject("System.Collections.ArrayList")
    'Backup worksheet
    ActiveSheet.Copy After:=ActiveSheet
    On Error GoTo ErrorHandler
    ActiveSheet.Name = BackupLabel & Format$(Now, DateFormat)
    On Error GoTo 0
    '@Ignore IndexedDefaultMemberAccess
    ActiveWorkbook.Sheets(1).Select
    'Iterating over all items' files in data file folder
    Dim File As Object
    For Each File In CreateObject("Scripting.FileSystemObject").GetFolder(DataFilePath).Files
        'Finding item's cell
        Dim Item As String
        Item = Library.GetFileNameWithoutExtension(File)
        Dim ItemCell As Range
        Set ItemCell = ActiveSheet.Columns(ItemColumn).Find(Item)
        'Processing item's data if cell is found, otherwise adding it to missing list
        If Not ItemCell Is Nothing Then
            Dim ItemRow As Long
            ItemRow = ItemCell.Row
            Dim ImportData() As String
            ImportData = Split(GetLastLine(File.Path)(0), Sep)
            'Accounting for kilo-unit
            Dim CurrentAmount As Double
            CurrentAmount = Replace(ImportData(ImportsCurrentAmountColumn), ImportUnit, vbNullString)
            If InStr(ActiveSheet.Cells(ItemRow, UnitColumn).Value, KiloUnitPrefix) > 0 Then
                CurrentAmount = CurrentAmount / 1000
            End If
            'Changing data in Excel table only if imported data is newer
            If CDate(ActiveSheet.Cells(ItemRow, LastChangedDateColumn).Value) < CDate(ImportData(ImportsLastChangedDateColumn)) Then
                'BB date
                ActiveSheet.Cells(ItemRow, BBDateColumn).Value = CDate(ImportData(ImportsCurrentBBDateColumn))
                'Last changed date
                ActiveSheet.Cells(ItemRow, LastChangedDateColumn).Value = Now
                'Amount
                Dim PreviousAmount As Double
                PreviousAmount = ActiveSheet.Cells(ItemRow, NewAmountColumn).Value
                ActiveSheet.Cells(ItemRow, PreviousAmountColum).Value = PreviousAmount
                Dim Diff As Double
                Diff = Math.Round(CurrentAmount - PreviousAmount, Decimals)
                ActiveSheet.Cells(ItemRow, AmountDiffColumn).Value = Diff
            End If
        ElseIf InStr(BlacklistedItems, Item) = 0 Then
            MissingItems.Add Item
        End If
    Next
    MsgBox (SuccessInfo)
    'Showing missing items list
    If MissingItems.Count > 0 Then
        Dim MissingItem As Variant
        Dim MissingItemsListString As String
        MissingItemsListString = EntryNotAvailableWarning
        For Each MissingItem In MissingItems
            MissingItemsListString = MissingItemsListString & MissingItem & vbNewLine
        Next
        MsgBox (MissingItemsListString)
    End If
    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then
        Err.Clear
        Application.DisplayAlerts = False
        ActiveWorkbook.ActiveSheet.Delete
        Application.DisplayAlerts = True
        '@Ignore IndexedDefaultMemberAccess
        ActiveWorkbook.Sheets(1).Select
        '@Ignore VariableNotUsed
        Dim Whatever As Variant
        Whatever = MsgBox(DoneAlreadyWarning, vbExclamation)
    End If
End Sub
