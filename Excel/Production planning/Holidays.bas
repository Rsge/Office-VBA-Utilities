Attribute VB_Name = "Holidays"
Option Explicit

Private Function ShowWeekends(CurrentDate As Date) As String
    Select Case Weekday(CurrentDate)
        Case vbSaturday, vbSunday
            ShowWeekends = "Weekend"
        Case Else
            ShowWeekends = ""
    End Select
End Function

Private Function ShowHolidays(CurrentDate As Date, Holidays As Range) As String
    If Holidays.Columns.Count <> 2 Then Error (1)

    Dim CurrentRow As Range
    Dim TempName As String
    Dim TempDate As Date
    For Each CurrentRow In Holidays.Rows
        TempName = CurrentRow.Cells(1, 1)
        TempDate = CurrentRow.Cells(1, 2)
        If CurrentDate = TempDate Then
            ShowHolidays = TempName
            Exit Function
        End If
    Next
    ShowHolidays = ""
End Function

Private Function ShowBridgingDays(CurrentDate As Date, BridgingDays As Range) As String
    If BridgingDays.Columns.Count <> 1 Then Error (1)
    
    Dim Cell As Range
    Dim TempDate As Date
    For Each Cell In BridgingDays.Cells
        TempDate = Cell
        If CurrentDate = TempDate Then
            ShowBridgingDays = "Bridge day"
            Exit Function
        End If
    Next
    ShowBridgingDays = ""
End Function

Private Function ShowCompanyHolidays(CurrentDate As Date, CompanyHolidays As Range) As String
    If CompanyHolidays.Columns.Count <> 2 Then Error (1)
    
    Dim CurrentRow As Range
    Dim FromDate As Date
    Dim ToDate As Date
    For Each CurrentRow In CompanyHolidays.Rows
        FromDate = CurrentRow.Cells(1, 1)
        ToDate = CurrentRow.Cells(1, 2)
        If FromDate <= CurrentDate And CurrentDate <= ToDate Then
            ShowCompanyHolidays = "Company holidays"
            Exit Function
        End If
    Next
    ShowCompanyHolidays = ""
End Function


Public Function ShowWorkFreeDays(CurrentDate As Date) As String
    With Worksheets("Holidays")
        ShowWorkFreeDays = ShowHolidays(CurrentDate, .Range("Holidays")) + ShowBridgingDays(CurrentDate, .Range("BridgeDays")) + ShowCompanyHolidays(CurrentDate, .Range("CompanyHolidays"))
    End With
    If Len(ShowWorkFreeDays) = 0 Then
        ShowWorkFreeDays = ShowWeekends(CurrentDate)
    End If
End Function

Public Function NoProduction(CurrentDate As Date) As Boolean
    NoProduction = ShowWorkFreeDays(CurrentDate) <> "" And ShowWorkFreeDays(CurrentDate) <> "Company holidays"
End Function

