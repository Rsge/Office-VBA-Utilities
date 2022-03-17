Attribute VB_Name = "Library"
Attribute VB_Description = "Module for general constants."
'@IgnoreModule InvalidAnnotation
'@Folder "Production planning"
'@ModuleDescription "Module for general constants."
Option Explicit

'Row constant
'@VariableDescription "Index of starting date's row."
Public Const StartingDateRow As Long = 2
'@VariableDescription "Index of first row with data."
Public Const StartingRow As Long = 5

'Column constants
'@VariableDescription "Index of starting date's column in data table."
Public Const StartingDateColumn As Long = 1
'@VariableDescription "Index of date values' column."
Public Const DateColumn As Long = 1
'@VariableDescription "Index of items' job numbers' column in data table."
Public Const JobColumn As Long = 2
'@VariableDescription "Index of items' column in data table."
Public Const ItemColumn As Long = 3
'@VariableDescription "Index of amounts' column in data table."
Public Const AmountColumn As Long = 4
'[AT]VariableDescription "Index of completed jobs' column in data table."
'Public Const CompletedJobsColumn As Long = 5
'[AT]VariableDescription "Index of due jobs' column in data table."
'Public Const DueJobsColumn As Long = 6
'@VariableDescription "Index of remaining capacities' column in data table."
Public Const RemainingCapacityColumn As Long = 7
'@VariableDescription "Index of holiday detection column in data table."
Public Const HolidaysColumn As Long = 8
'@VariableDescription "Index of special slowdowns' column in data table."
Public Const SlowdownsColumn As Long = 9
'@VariableDescription "Index of jobs' definitions' column in data worksheet."
Public Const JobsDefColumn As Long = 11
'@VariableDescription "Index of jobs' due dates' definitions' column in data worksheet."
Public Const JobsDueDatesColumn As Long = 12
'@VariableDescription "Ascii value to add to column index to get column letter."
Public Const ColumnLetterAscii As Long = 64

'String constants
'@VariableDescription "Colon symbol."
Public Const Colon As String = ":"
'@VariableDescription "Comma symbol with space."
Public Const Comma As String = ", "
