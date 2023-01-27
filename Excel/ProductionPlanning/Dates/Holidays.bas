Attribute VB_Name = "Holidays"
Attribute VB_Description = "Finding of holidays and other non-productive dates."
'@Folder("ProductionPlanning.Dates")
'@ModuleDescription("Finding of holidays and other non-productive dates.")
Option Explicit

' String constants
'@VariableDescription("Label for weekends.")
Private Const m_weekendLabel As String = "Weekend"
Attribute m_weekendLabel.VB_VarDescription = "Label for weekends."
'@VariableDescription("Label for bridging days (Days in between holidays and weekends).")
Private Const m_bridgingDayLabel As String = "Bridging day"
Attribute m_bridgingDayLabel.VB_VarDescription = "Label for bridging days (Days in between holidays and weekends)."
'@VariableDescription("Label for company-wide holidays.")
Private Const m_companyHolidaysLabel As String = "Company holidays"
Attribute m_companyHolidaysLabel.VB_VarDescription = "Label for company-wide holidays."
'@VariableDescription("Label for legal holidays.")
Private Const m_holidaysWorksheetName As String = "Holidays"
Attribute m_holidaysWorksheetName.VB_VarDescription = "Label for legal holidays."
'@VariableDescription("Name of table containing the holidays.")
Private Const m_holidaysTableName As String = "Holidays"
Attribute m_holidaysTableName.VB_VarDescription = "Name of table containing the holidays."
'@VariableDescription("Name of table containing the briding days.")
Private Const m_bridgingDaysTableName As String = "BridgingDays"
Attribute m_bridgingDaysTableName.VB_VarDescription = "Name of table containing the briding days."
'@VariableDescription("Name of table containing company holidays.")
Private Const m_companyHolidaysTableName As String = "CompanyHolidays"
Attribute m_companyHolidaysTableName.VB_VarDescription = "Name of table containing company holidays."
' Column constants
'@VariableDescription("Index of holidays' names' column.")
Private Const m_holidayNameColumn As Long = 1
Attribute m_holidayNameColumn.VB_VarDescription = "Index of holidays' names' column."
'@VariableDescription("Index of holidays' dates' column.")
Private Const m_holidayDateColumn As Long = 2
Attribute m_holidayDateColumn.VB_VarDescription = "Index of holidays' dates' column."

' ————————————————————————————————————————————————————— '


'@Description("Shows if a given date is a weekend.")
Private Function ShowWeekends(ByVal currentDate As Date) As String
Attribute ShowWeekends.VB_Description = "Shows if a given date is a weekend."
    Select Case Weekday(currentDate)
        Case vbSaturday, vbSunday
            ShowWeekends = m_weekendLabel
        Case Else
            ShowWeekends = vbNullString
    End Select
End Function

'@Description("Shows if a given date is a legal holiday according to an extra table containing all legal holidays.")
Private Function ShowHolidays(ByVal currentDate As Date, ByVal holidays As Range) As String
Attribute ShowHolidays.VB_Description = "Shows if a given date is a legal holiday according to an extra table containing all legal holidays."
    If holidays.Columns.Count <> 2 Then Err.Raise (1)
    Dim currentRow As Range
    Dim holidayName As String
    Dim holidayDate As Date
    For Each currentRow In holidays.Rows
        holidayName = currentRow.Cells.Item(1, m_holidayNameColumn)
        holidayDate = currentRow.Cells.Item(1, m_holidayDateColumn)
        If currentDate = holidayDate Then
            ShowHolidays = holidayName
            Exit Function
        End If
    Next
    ShowHolidays = vbNullString
End Function

'@Description("Shows if a given date is a bridging day according to a corresponding table.")
Private Function ShowBridgingDays(ByVal currentDate As Date, ByVal bridgingDays As Range) As String
Attribute ShowBridgingDays.VB_Description = "Shows if a given date is a bridging day according to a corresponding table."
    If bridgingDays.Columns.Count <> 1 Then Err.Raise (1)
    Dim cell As Range
    Dim bridgingDayDate As Date
    For Each cell In bridgingDays.Cells
        bridgingDayDate = cell.Value
        If currentDate = bridgingDayDate Then
            ShowBridgingDays = m_bridgingDayLabel
            Exit Function
        End If
    Next
    ShowBridgingDays = vbNullString
End Function

'@Description("Shows if a given date is a company holiday according to a corresponding table.")
Private Function ShowCompanyHolidays(ByVal currentDate As Date, ByVal companyHolidays As Range) As String
Attribute ShowCompanyHolidays.VB_Description = "Shows if a given date is a company holiday according to a corresponding table."
    If companyHolidays.Columns.Count <> 2 Then Err.Raise (1)
    Dim currentRow As Range
    Dim fromDate As Date
    Dim toDate As Date
    For Each currentRow In companyHolidays.Rows
        fromDate = currentRow.Cells.Item(1, m_holidayNameColumn)
        toDate = currentRow.Cells.Item(1, m_holidayDateColumn)
        If fromDate <= currentDate And currentDate <= toDate Then
            ShowCompanyHolidays = m_companyHolidaysLabel
            Exit Function
        End If
    Next
    ShowCompanyHolidays = vbNullString
End Function

' ————————————————————————————————————————————————————— '

'@Description("Shows all days which are completely work-free.")
Public Function ShowWorkFreeDays(ByVal currentDate As Date) As String
Attribute ShowWorkFreeDays.VB_Description = "Shows all days which are completely work-free."
    With ActiveWorkbook.Worksheets.Item(m_holidaysWorksheetName)
        ShowWorkFreeDays = ShowHolidays(currentDate, .Range(m_holidaysTableName)) + ShowBridgingDays(currentDate, .Range(m_bridgingDaysTableName)) + ShowCompanyHolidays(currentDate, .Range(m_companyHolidaysTableName))
    End With
    If LenB(ShowWorkFreeDays) = 0 Then
        ShowWorkFreeDays = ShowWeekends(currentDate)
    End If
End Function

'@Description("Shows all days on which there is no production.")
Public Function NoProduction(ByVal currentDate As Date) As Boolean
Attribute NoProduction.VB_Description = "Shows all days on which there is no production."
    NoProduction = ShowWorkFreeDays(currentDate) <> vbNullString And ShowWorkFreeDays(currentDate) <> m_companyHolidaysLabel
End Function
