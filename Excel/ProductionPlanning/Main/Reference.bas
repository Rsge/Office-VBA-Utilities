Attribute VB_Name = "Reference"
Attribute VB_Description = "General constants."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("General constants.")
Option Explicit

' Row constants
'@VariableDescription("Index of starting date's row.")
Public Const StartingDateRow As Long = 2
Attribute StartingDateRow.VB_VarDescription = "Index of starting date's row."
'@VariableDescription("Index of first row with data.")
Public Const StartingRow As Long = 5
Attribute StartingRow.VB_VarDescription = "Index of first row with data."

' Column constants
'@VariableDescription("Index of starting date's column in data table.")
Public Const StartingDateColumn As Long = 1
Attribute StartingDateColumn.VB_VarDescription = "Index of starting date's column in data table."
'@VariableDescription("Index of date values' column.")
Public Const DateColumn As Long = 1
Attribute DateColumn.VB_VarDescription = "Index of date values' column."
'@VariableDescription("Index of items' job numbers' column in data table.")
Public Const JobColumn As Long = 2
Attribute JobColumn.VB_VarDescription = "Index of items' job numbers' column in data table."
'@VariableDescription("Index of items' column in data table.")
Public Const ItemColumn As Long = 3
Attribute ItemColumn.VB_VarDescription = "Index of items' column in data table."
'@VariableDescription("Index of amounts' column in data table.")
Public Const AmountColumn As Long = 4
Attribute AmountColumn.VB_VarDescription = "Index of amounts' column in data table."
''@VariableDescription("Index of completed jobs' column in data table.")
'Public Const CompletedJobsColumn As Long = 5
''@VariableDescription("Index of due jobs' column in data table.")
'Public Const DueJobsColumn As Long = 6
'@VariableDescription("Index of remaining capacities' column in data table.")
Public Const RemainingCapacityColumn As Long = 7
Attribute RemainingCapacityColumn.VB_VarDescription = "Index of remaining capacities' column in data table."
'@VariableDescription("Index of holiday detection column in data table.")
Public Const HolidaysColumn As Long = 8
Attribute HolidaysColumn.VB_VarDescription = "Index of holiday detection column in data table."
'@VariableDescription("Index of special slowdowns' column in data table.")
Public Const SlowdownsColumn As Long = 9
Attribute SlowdownsColumn.VB_VarDescription = "Index of special slowdowns' column in data table."
'@VariableDescription("Index of jobs' definitions' column in data worksheet.")
Public Const JobsDefColumn As Long = 11
Attribute JobsDefColumn.VB_VarDescription = "Index of jobs' definitions' column in data worksheet."
'@VariableDescription("Index of jobs' due dates' definitions' column in data worksheet.")
Public Const JobsDueDatesColumn As Long = 12
Attribute JobsDueDatesColumn.VB_VarDescription = "Index of jobs' due dates' definitions' column in data worksheet."
'@VariableDescription("Ascii value to add to column index to get column letter.")
Public Const ColumnLetterAscii As Long = 64
Attribute ColumnLetterAscii.VB_VarDescription = "Ascii value to add to column index to get column letter."

' String constants
'@VariableDescription("Colon symbol.")
Public Const Colon As String = ":"
Attribute Colon.VB_VarDescription = "Colon symbol."
'@VariableDescription("Comma symbol with space.")
Public Const Comma As String = ", "
Attribute Comma.VB_VarDescription = "Comma symbol with space."

