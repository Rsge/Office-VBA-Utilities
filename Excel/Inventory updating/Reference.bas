Attribute VB_Name = "Reference"
Attribute VB_Description = "General constants."
'@Folder "Inventory updating"
'@ModuleDescription "General constants."
Option Explicit

'Basic constants
'@VariableDescription "Extension of data files."
Public Const Ext As String = "*.csv"
Attribute Ext.VB_VarDescription = "Extension of data files."
'@VariableDescription "Separator character in data file."
Public Const Sep As String = ";"
Attribute Sep.VB_VarDescription = "Separator character in data file."
''@VariableDescription "Colon symbol."
'Public Const Colon As String = ":"
'@VariableDescription "Number of decimal places to use in rounding."
Public Const Decimals As Long = 3
Attribute Decimals.VB_VarDescription = "Number of decimal places to use in rounding."
'@VariableDescription "Unit string used in data files."
Public Const ImportUnit As String = " g"
Attribute ImportUnit.VB_VarDescription = "Unit string used in data files."
'@VariableDescription "'Thousands' prefix of unit."
Public Const KiloUnitPrefix As String = "k"
Attribute KiloUnitPrefix.VB_VarDescription = "'Thousands' prefix of unit."
'@VariableDescription "Format of date in Excel table."
Public Const DateFormat As String = "dd.mm.yy"
Attribute DateFormat.VB_VarDescription = "Format of date in Excel table."
'@VariableDescription "Label for worksheet backup."
Public Const BackupLabel As String = "Backup "
Attribute BackupLabel.VB_VarDescription = "Label for worksheet backup."
'@VariableDescription "Item blacklisted from being processed."
Public Const BlacklistedItem As String = "WATER;12345"
Attribute BlacklistedItem.VB_VarDescription = "Item blacklisted from being processed."

'Row constants
'@VariableDescription "Index of row of cell containing data file path."
Public Const PathCellRow As Long = 2
Attribute PathCellRow.VB_VarDescription = "Index of row of cell containing data file path."
''@VariableDescription "Index of row with first data."
'Public Const StartingRow As Long = 3

'Excel table column constants (1-based)
'@VariableDescription "Index of items' column in Excel table."
Public Const ItemColumn As Long = 1
Attribute ItemColumn.VB_VarDescription = "Index of items' column in Excel table."
''@VariableDescription "Index of descriptions' column in Excel table."
'Public Const DescriptionColumn As Long = 2
'@VariableDescription "Index of BB dates' column in Excel table."
Public Const BBDateColumn As Long = 3
Attribute BBDateColumn.VB_VarDescription = "Index of BB dates' column in Excel table."
'@VariableDescription "Index of amounts' units' column in Excel table."
Public Const UnitColumn As Long = 4
Attribute UnitColumn.VB_VarDescription = "Index of amounts' units' column in Excel table."
'@VariableDescription "Index of previous amounts' column in Excel table."
Public Const PreviousAmountColum As Long = 5
Attribute PreviousAmountColum.VB_VarDescription = "Index of previous amounts' column in Excel table."
'@VariableDescription "Index of amount differences' column in Excel table."
Public Const AmountDiffColumn As Long = 6
Attribute AmountDiffColumn.VB_VarDescription = "Index of amount differences' column in Excel table."
'@VariableDescription "Index of the new amounts' column in Excel table."
Public Const NewAmountColumn As Long = 7
Attribute NewAmountColumn.VB_VarDescription = "Index of the new amounts' column in Excel table."
'@VariableDescription "Index of last changes' dates' column in Excel table."
Public Const LastChangedDateColumn As Long = 8
Attribute LastChangedDateColumn.VB_VarDescription = "Index of last changes' dates' column in Excel table."
'@VariableDescription "Index of column of cell with files' path in Excel table."
Public Const PathCellColumn As Long = 14
Attribute PathCellColumn.VB_VarDescription = "Index of column of cell with files' path in Excel table."
''@VariableDescription "Ascii value to add to column index to get column letter."
'Public Const ColumnLetterAscii As Long = 64

'CSV data file column constants (0-based)
'@VariableDescription "Index of last changed dates' column in CSV data file."
Public Const ImportsLastChangedDateColumn As Long = 1
Attribute ImportsLastChangedDateColumn.VB_VarDescription = "Index of last changed dates' column in CSV data file."
'@VariableDescription "Index of current BB dates' column in CSV data file."
Public Const ImportsCurrentBBDateColumn As Long = 6
Attribute ImportsCurrentBBDateColumn.VB_VarDescription = "Index of current BB dates' column in CSV data file."
'@VariableDescription "Index of current amounts' column in CSV data file."
Public Const ImportsCurrentAmountColumn As Long = 8
Attribute ImportsCurrentAmountColumn.VB_VarDescription = "Index of current amounts' column in CSV data file."
