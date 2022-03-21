Attribute VB_Name = "Holidays"
Attribute VB_Description = "Finding of holidays and other non-productive dates."
'@Folder "Production planning"
'@ModuleDescription "Finding of holidays and other non-productive dates."
Option Explicit

'String constants
'@VariableDescription "Label for weekends."
Private Const WeekendLabel As String = "Weekend"
Attribute WeekendLabel.VB_VarDescription = "Label for weekends."
'@VariableDescription "Label for bridging days (Days in between holidays and weekends)."
Private Const BridgingDayLabel As String = "Bridging day"
Attribute BridgingDayLabel.VB_VarDescription = "Label for bridging days (Days in between holidays and weekends)."
'@VariableDescription "Label for company-wide holidays."
Private Const CompanyHolidaysLabel As String = "Company holidays"
Attribute CompanyHolidaysLabel.VB_VarDescription = "Label for company-wide holidays."
'@VariableDescription "Label for legal holidays."
Private Const HolidaysWorksheetName As String = "Holidays"
Attribute HolidaysWorksheetName.VB_VarDescription = "Label for legal holidays."
'@VariableDescription "Name of table containing the holidays."
Private Const HolidaysTableName As String = "Holidays"
Attribute HolidaysTableName.VB_VarDescription = "Name of table containing the holidays."
'@VariableDescription "Name of table containing the briding days."
Private Const BridgingDaysTableName As String = "BridgingDays"
Attribute BridgingDaysTableName.VB_VarDescription = "Name of table containing the briding days."
'@VariableDescription "Name of table containing company holidays."
Private Const CompanyHolidaysTableName As String = "CompanyHolidays"
Attribute CompanyHolidaysTableName.VB_VarDescription = "Name of table containing company holidays."
'Column constants
'@VariableDescription "Index of holidays' names' column."
Private Const HolidayNameColumn As Long = 1
Attribute HolidayNameColumn.VB_VarDescription = "Index of holidays' names' column."
'@VariableDescription "Index of holidays' dates' column."
Private Const HolidayDateColumn As Long = 2
Attribute HolidayDateColumn.VB_VarDescription = "Index of holidays' dates' column."


'@Description "Shows if a given date is a weekend."
Private Function ShowWeekends(ByVal CurrentDate As Date) As String
Attribute ShowWeekends.VB_Description = "Shows if a given date is a weekend."
    Select Case Weekday(CurrentDate)
        Case vbSaturday, vbSunday
            ShowWeekends = WeekendLabel
        Case Else
            ShowWeekends = vbNullString
    End Select
End Function

'@Description "Shows if a given date is a legal holiday according to an extra table containing all legal holidays."
Private Function ShowHolidays(ByVal CurrentDate As Date, ByVal Holidays As Range) As String
Attribute ShowHolidays.VB_Description = "Shows if a given date is a legal holiday according to an extra table containing all legal holidays."
    If Holidays.Columns.Count <> 2 Then Err.Raise (1)
    Dim CurrentRow As Range
    Dim HolidayName As String
    Dim HolidayDate As Date
    For Each CurrentRow In Holidays.Rows
        HolidayName = CurrentRow.Cells.Item(1, HolidayNameColumn)
        HolidayDate = CurrentRow.Cells.Item(1, HolidayDateColumn)
        If CurrentDate = HolidayDate Then
            ShowHolidays = HolidayName
            Exit Function
        End If
    Next
    ShowHolidays = vbNullString
End Function

'@Description "Shows if a given date is a bridging day according to a corresponding table."
Private Function ShowBridgingDays(ByVal CurrentDate As Date, ByVal BridgingDays As Range) As String
Attribute ShowBridgingDays.VB_Description = "Shows if a given date is a bridging day according to a corresponding table."
    If BridgingDays.Columns.Count <> 1 Then Err.Raise (1)
    Dim Cell As Range
    Dim BridgingDayDate As Date
    For Each Cell In BridgingDays.Cells
        BridgingDayDate = Cell.Value
        If CurrentDate = BridgingDayDate Then
            ShowBridgingDays = BridgingDayLabel
            Exit Function
        End If
    Next
    ShowBridgingDays = vbNullString
End Function

'@Description "Shows if a given date is a company holiday according to a corresponding table."
Private Function ShowCompanyHolidays(ByVal CurrentDate As Date, ByVal CompanyHolidays As Range) As String
Attribute ShowCompanyHolidays.VB_Description = "Shows if a given date is a company holiday according to a corresponding table."
    If CompanyHolidays.Columns.Count <> 2 Then Err.Raise (1)
    Dim CurrentRow As Range
    Dim FromDate As Date
    Dim ToDate As Date
    For Each CurrentRow In CompanyHolidays.Rows
        FromDate = CurrentRow.Cells.Item(1, HolidayNameColumn)
        ToDate = CurrentRow.Cells.Item(1, HolidayDateColumn)
        If FromDate <= CurrentDate And CurrentDate <= ToDate Then
            ShowCompanyHolidays = CompanyHolidaysLabel
            Exit Function
        End If
    Next
    ShowCompanyHolidays = vbNullString
End Function


'@Description "Shows all days which are completely work-free."
Public Function ShowWorkFreeDays(ByVal CurrentDate As Date) As String
Attribute ShowWorkFreeDays.VB_Description = "Shows all days which are completely work-free."
    '@Ignore IndexedDefaultMemberAccess
    With ActiveWorkbook.Worksheets(HolidaysWorksheetName)
        ShowWorkFreeDays = ShowHolidays(CurrentDate, .Range(HolidaysTableName)) + ShowBridgingDays(CurrentDate, .Range(BridgingDaysTableName)) + ShowCompanyHolidays(CurrentDate, .Range(CompanyHolidaysTableName))
    End With
    If LenB(ShowWorkFreeDays) = 0 Then
        ShowWorkFreeDays = ShowWeekends(CurrentDate)
    End If
End Function

'@Description "Shows all days on which there is no production."
Public Function NoProduction(ByVal CurrentDate As Date) As Boolean
Attribute NoProduction.VB_Description = "Shows all days on which there is no production."
    NoProduction = ShowWorkFreeDays(CurrentDate) <> vbNullString And ShowWorkFreeDays(CurrentDate) <> CompanyHolidaysLabel
End Function
