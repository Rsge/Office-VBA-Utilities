Attribute VB_Name = "DataReset"
Attribute VB_Description = "Reset imported weight data to backup point."
'@Folder("InventoryUpdating.Imports")
'@ModuleDescription("Reset imported weight data to backup point.")
Option Explicit

'@EntryPoint
'@Description("Resets imported weight data to backup point.")
Public Sub ResetTable()
Attribute ResetTable.VB_Description = "Resets imported weight data to backup point."
    ' If there's a backup to "load"...
    If ContainsSheetStartingWith(BackupSheetLabel) Then
        ' Reset to backup sheet.
        ActiveWorkbook.Worksheets.Item(1).Select
        Dim name_ As String
        name_ = ActiveWorkbook.ActiveSheet.Name
        DeleteSheetsStartingWith name_
        ActiveWorkbook.ActiveSheet.Name = name_
        ' Mark already happened import.
        GetActCell(ImportPathAndResetMarkerRow, ResetMarkerColumn).Value = ResetMarkerMsg
    Else
        WarnBox NoResetWarning
    End If
End Sub
