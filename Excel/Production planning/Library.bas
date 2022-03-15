Attribute VB_Name = "Library"
Attribute VB_Description = "Module for general constants."
'@Folder "Production planning"
'@ModuleDescription "Module for general constants."
Option Explicit

'Row constant
Public Const StartingDateRow As Long = 2
Public Const StartingRow As Long = 5

'Column constants
Public Const StartingDateColumn As Long = 1
Public Const DateColumn As Long = 1
Public Const JobColumn As Long = 2
Public Const ItemColumn As Long = 3
Public Const AmountColumn As Long = 4
'Public Const CompletedJobsColumn As Long = 5
'Public Const DueJobsColumn As Long = 6
Public Const RemainingCapacityColumn As Long = 7
Public Const HolidaysColumn As Long = 8
Public Const SlowdownsColumn As Long = 9
Public Const JobsDefColumn As Long = 11
Public Const JobsDueDatesColumn As Long = 12
Public Const ColumnLetterAscii As Long = 64

'String constants
Public Const Colon As String = ":"
Public Const Comma As String = ", "
