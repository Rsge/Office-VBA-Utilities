Attribute VB_Name = "Reference"
Attribute VB_Description = "General constants."
'@Folder("InventoryUpdating.Base")
'@ModuleDescription("General constants.")
Option Explicit

' Feature toggles
'@VariableDescription("Create a copy of the workbook and do the processing in that copy.")
Public Const CreateWBCopy As Boolean = True
Attribute CreateWBCopy.VB_VarDescription = "Create a copy of the workbook and do the processing in that copy."

' Basic constants
'@VariableDescription("Separator character in data file.")
Public Const Sep As String = ";"
Attribute Sep.VB_VarDescription = "Separator character in data file."
'@VariableDescription("Search pattern for import data files.")
Public Const DataFilePattern As String = "*.csv"
Attribute DataFilePattern.VB_VarDescription = "Search pattern for import data files."
'@VariableDescription("Number of decimal places to use in rounding.")
Public Const Decimals As Long = 3
Attribute Decimals.VB_VarDescription = "Number of decimal places to use in rounding."
'@VariableDescription("Unit string used in data files.")
Public Const ImportUnit As String = " g"
Attribute ImportUnit.VB_VarDescription = "Unit string used in data files."
'@VariableDescription("'Thousands' prefix of unit.")
Public Const KiloUnitPrefix As String = "k"
Attribute KiloUnitPrefix.VB_VarDescription = "'Thousands' prefix of unit."
'@VariableDescription("Unit of litres, prompting division by 1000.")
Public Const LitersUnit As String = "l"
Attribute LitersUnit.VB_VarDescription = "Unit of litres, prompting division by 1000."
'@VariableDescription("Regex-format of file name. Whats in brackets will be constant, the rest varies depending on which file is looked at.")
Public Const FileNameFormatRegex As String = "(\w+)\s?\d*[\.-/]?\d*[\.-/]?\d*"
Attribute FileNameFormatRegex.VB_VarDescription = "Regex-format of file name. Whats in brackets will be constant, the rest varies depending on which file is looked at."

' - Localized strings -
' General
'@VariableDescription("Title for warning MsgBox.")
Public Const WarnBoxTitle As String = "Warning!"
Attribute WarnBoxTitle.VB_VarDescription = "Title for warning MsgBox."
'@VariableDescription("Label for worksheet backup.")
Public Const BackupSheetLabel As String = "Backup "
Attribute BackupSheetLabel.VB_VarDescription = "Label for worksheet backup."
'@VariableDescription("Name of definitions worksheet.")
Public Const DefSheetName As String = "Definitions"
Attribute DefSheetName.VB_VarDescription = "Name of definitions worksheet."
'@VariableDescription("Name of special items table.")
Public Const SpecialItemsTableName As String = "SpecialItems"
Attribute SpecialItemsTableName.VB_VarDescription = "Name of special items table."
'@VariableDescription("Marker at end of file name of item(s) with a special variant.")
Public Const SpecialItemFileMarker As String = "S"
Attribute SpecialItemFileMarker.VB_VarDescription = "Marker at end of file name of item(s) with a special variant."
'@VariableDescription("Marker starting the description of item(s) with a special variant.")
Public Const SpecialItemDescriptionMarker As String = "Special"
Attribute SpecialItemDescriptionMarker.VB_VarDescription = "Marker starting the description of item(s) with a special variant."
'@VariableDescription("Name of blacklisted items table.")
Public Const BlacklistedItemsTableName As String = "Blacklist"
Attribute BlacklistedItemsTableName.VB_VarDescription = "Name of blacklisted items table."
'@VariableDescription("Format of date in Excel table.")
Public Const DataDateFormat As String = "dd/mm/yy;@"
Attribute DataDateFormat.VB_VarDescription = "Format of date in Excel table."
'@VariableDescription("Format of date in working workbook's file name.")
Public Const ActFileDateFormat As String = " yyyy-mm-dd"
Attribute ActFileDateFormat.VB_VarDescription = "Format of date in working workbook's file name."
'@VariableDescription("Format of date in export workbook's file name.")
Public Const ExportDateFormat As String = " yyyy"
Attribute ExportDateFormat.VB_VarDescription = "Format of date in export workbook's file name."
'@VariableDescription("Placeholder date in file data for non-existent BB-date.")
Public Const PlaceholderDate As String = "11.11.1111"
Attribute PlaceholderDate.VB_VarDescription = "Placeholder date in file data for non-existent BB-date."

' Library
'@VariableDescription("Warning if no save location specified.")
Public Const NoPathWarning As String = "No path for {} configured" & vbNewLine _
                                     & "or path doesn't exist." & vbNewLine _
                                     & "Please specify a path."
Attribute NoPathWarning.VB_VarDescription = "Warning if no save location specified."
'@VariableDescription("Warning if no files found at save location.")
Public Const NoFilesWarning As String = "No files found for {} at specified location." & vbNewLine _
                                      & "Please specify a different folder or abort and add files first."
Attribute NoFilesWarning.VB_VarDescription = "Warning if no files found at save location."

' DataImport
'@VariableDescription("Descriptive label of the import files.")
Public Const ImportLabel As String = "import data"
Attribute ImportLabel.VB_VarDescription = "Descriptive label of the import files."
'@VariableDescription("Descriptive label of the export file.")
Public Const ExportLabel As String = "the export workbook"
Attribute ExportLabel.VB_VarDescription = "Descriptive label of the export file."
'@VariableDescription("Warning for export file being read-only.")
Public Const ReadOnlyWarning As String = "The export workbook is read-only at the moment." & vbNewLine & _
                                         "Please ensure the workbook has been closed everywhere and try again."
Attribute ReadOnlyWarning.VB_VarDescription = "Warning for export file being read-only."
'@VariableDescription("Warning for a file's item number not being present in table.")
Public Const EntryNotAvailableWarning As String = "No entry exists for the following items." & vbNewLine _
                                                & "Please add the correct description and check the unit " _
                                                & "in this and the export file open in the background " _
                                                & "or fix the import data files." & vbNewLine
Attribute EntryNotAvailableWarning.VB_VarDescription = "Warning for a file's item number not being present in table."
'@VariableDescription("Info about successful processing.")
Public Const SuccessInfo As String = "Data processed successfully."
Attribute SuccessInfo.VB_VarDescription = "Info about successful processing."
'@VariableDescription("Warning about processing already having been done.")
Public Const DoneAlreadyWarning As String = "Data import was already carried out today." & vbNewLine _
                                          & "Use the Reset button to revert today's processing."
Attribute DoneAlreadyWarning.VB_VarDescription = "Warning about processing already having been done."

' DataReset
'@VariableDescription("Message in cell used as marker for reset.")
Public Const ResetMarkerMsg As String = "[DO NOT DELETE YOURSELF]"
Attribute ResetMarkerMsg.VB_VarDescription = "Message in cell used as marker for reset."
'@VariableDescription("Warning about reset not being possible because no backup is available.")
Public Const NoResetWarning As String = "There is no backup to revert to."
Attribute NoResetWarning.VB_VarDescription = "Warning about reset not being possible because no backup is available."

' - Region constants -
' Excel table range string constants
'@VariableDescription("Index of row of cell containing import data files' path.")
Public Const DataRegionStartCell As String = "A1"
Attribute DataRegionStartCell.VB_VarDescription = "Index of row of cell containing import data files' path."

' Excel table row constants (First = 1)
'@VariableDescription("Index of row of cell containing import data files' path.")
Public Const ImportPathAndResetMarkerRow As Long = 1
Attribute ImportPathAndResetMarkerRow.VB_VarDescription = "Index of row of cell containing import data files' path."
'@VariableDescription("Index of row of cell containing export workbook path.")
Public Const ExportPathRow As Long = 2
Attribute ExportPathRow.VB_VarDescription = "Index of row of cell containing export workbook path."
'@VariableDescription("Index of row with first data.")
Public Const StartingRow As Long = 3
Attribute StartingRow.VB_VarDescription = "Index of row with first data."

' Excel table column constants (First = 1)
'@VariableDescription("Index of items' column in Excel table.")
Public Const ItemColumn As Long = 1
Attribute ItemColumn.VB_VarDescription = "Index of items' column in Excel table."
'@VariableDescription("Index of descriptions' column in Excel table.")
Public Const DescriptionColumn As Long = 2
Attribute DescriptionColumn.VB_VarDescription = "Index of descriptions' column in Excel table."
'@VariableDescription("Index of BB dates' column in Excel table.")
Public Const BBDateColumn As Long = 3
Attribute BBDateColumn.VB_VarDescription = "Index of BB dates' column in Excel table."
'@VariableDescription("Index of amounts' units' column in Excel table.")
Public Const UnitColumn As Long = 4
Attribute UnitColumn.VB_VarDescription = "Index of amounts' units' column in Excel table."
'@VariableDescription("Index of previous amounts' column in Excel table.")
Public Const PreviousAmountColum As Long = 5
Attribute PreviousAmountColum.VB_VarDescription = "Index of previous amounts' column in Excel table."
'@VariableDescription("Index of amount differences' column in Excel table.")
Public Const AmountDiffColumn As Long = 6
Attribute AmountDiffColumn.VB_VarDescription = "Index of amount differences' column in Excel table."
'@VariableDescription("Index of the new amounts' column in Excel table.")
Public Const NewAmountColumn As Long = 7
Attribute NewAmountColumn.VB_VarDescription = "Index of the new amounts' column in Excel table."
'@VariableDescription("Index of last changes' dates' column in Excel table.")
Public Const LastChangedDateColumn As Long = 8
Attribute LastChangedDateColumn.VB_VarDescription = "Index of last changes' dates' column in Excel table."
'@VariableDescription("Index of reset marker cell's column in Excel table.")
Public Const ResetMarkerColumn As Long = 11
Attribute ResetMarkerColumn.VB_VarDescription = "Index of reset marker cell's column in Excel table."
'@VariableDescription("Index of column of cells with import data files' path in Excel table.")
Public Const PathCellsColumn As Long = 15
Attribute PathCellsColumn.VB_VarDescription = "Index of column of cells with import data files' path in Excel table."

' CSV data file column constants (First = 0)
'@VariableDescription("Index of last changed dates' column in CSV data file.")
Public Const ImportsLastChangedDateColumn As Long = 1
Attribute ImportsLastChangedDateColumn.VB_VarDescription = "Index of last changed dates' column in CSV data file."
'@VariableDescription("Index of current BB dates' column in CSV data file.")
Public Const ImportsCurrentBBDateColumn As Long = 6
Attribute ImportsCurrentBBDateColumn.VB_VarDescription = "Index of current BB dates' column in CSV data file."
'@VariableDescription("Index of current amounts' column in CSV data file.")
Public Const ImportsCurrentAmountColumn As Long = 8
Attribute ImportsCurrentAmountColumn.VB_VarDescription = "Index of current amounts' column in CSV data file."
