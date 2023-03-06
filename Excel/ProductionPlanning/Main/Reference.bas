Attribute VB_Name = "Reference"
Attribute VB_Description = "General constants."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("General constants.")
Option Explicit

' - Localized strings -
' Jobs
'@VariableDescription("Info string added to job on last day to indicate it'll finish in the future, beyond the current table's scopes.")
Public Const FutureInfo As String = "Future:"
Attribute FutureInfo.VB_VarDescription = "Info string added to job on last day to indicate it'll finish in the future, beyond the current table's scopes."
' Holidays
'@VariableDescription("Label for weekends.")
Public Const WeekendLabel As String = "Weekend"
Attribute WeekendLabel.VB_VarDescription = "Label for weekends."
'@VariableDescription("Label for bridging days (Days in between holidays and weekends).")
Public Const BridgingDayLabel As String = "Bridging day"
Attribute BridgingDayLabel.VB_VarDescription = "Label for bridging days (Days in between holidays and weekends)."
'@VariableDescription("Label for company-wide holidays.")
Public Const CompanyHolidaysLabel As String = "Company holidays"
Attribute CompanyHolidaysLabel.VB_VarDescription = "Label for company-wide holidays."
'@VariableDescription("Label for legal holidays.")
Public Const HolidaysWorksheetName As String = "Holidays"
Attribute HolidaysWorksheetName.VB_VarDescription = "Label for legal holidays."
'@VariableDescription("Name of table containing the holidays.")
Public Const HolidaysTableName As String = "Holidays"
Attribute HolidaysTableName.VB_VarDescription = "Name of table containing the holidays."
'@VariableDescription("Name of table containing the briding days.")
Public Const BridgingDaysTableName As String = "BridgingDays"
Attribute BridgingDaysTableName.VB_VarDescription = "Name of table containing the briding days."
'@VariableDescription("Name of table containing company holidays.")
Public Const CompanyHolidaysTableName As String = "CompanyHolidays"
Attribute CompanyHolidaysTableName.VB_VarDescription = "Name of table containing company holidays."
' Basic Utilities
'@VariableDescription("Question about up to which date the calculations should be cleared.")
Public Const DeletionQuestion As String = "Up to which date shall be deleted?" & vbNewLine & "(DD.MM.YYYY)"
Attribute DeletionQuestion.VB_VarDescription = "Question about up to which date the calculations should be cleared."
'@VariableDescription("Warning for input not being processable as a date.")
Public Const NoDateWarning As String = "Input can't be processed as a date." & vbNewLine & vbNewLine & DeletionQuestion
Attribute NoDateWarning.VB_VarDescription = "Warning for input not being processable as a date."
'@VariableDescription("Warning to check special slowdown after making changes to dates etc.")
Public Const SlowdownChangeWarning As String = "Please check special slowdown!"
Attribute SlowdownChangeWarning.VB_VarDescription = "Warning to check special slowdown after making changes to dates etc."
'@VariableDescription("Title of input box to show it needs an input.")
Public Const InputLabel As String = "Input"
Attribute InputLabel.VB_VarDescription = "Title of input box to show it needs an input."
'@VariableDescription("Title of MsgBox to show it contains a warning.")
Public Const WarningLabel As String = "Warning!"
Attribute WarningLabel.VB_VarDescription = "Title of MsgBox to show it contains a warning."
'@VariableDescription("Message for lifted worksheet protection.")
Public Const ProtectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Attribute ProtectionLifted.VB_VarDescription = "Message for lifted worksheet protection."
'@VariableDescription("Message for enforced worksheet protection.")
Public Const ProtectionEnabled As String = "Protection reestablished."
Attribute ProtectionEnabled.VB_VarDescription = "Message for enforced worksheet protection."
' Legacy
'@VariableDescription("Prefix for a date string to symbolize something happens on this day.)
Public Const DatePrefix As String = "On "

' String constants
'@VariableDescription("Colon symbol.")
Public Const Colon As String = ":"
Attribute Colon.VB_VarDescription = "Colon symbol."
'@VariableDescription("Comma symbol with space.")
Public Const Comma As String = ", "
Attribute Comma.VB_VarDescription = "Comma symbol with space."

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
