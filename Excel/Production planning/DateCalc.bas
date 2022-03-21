Attribute VB_Name = "DateCalc"
Attribute VB_Description = "Date calculation."
'@Folder "Production planning"
'@ModuleDescription "Date calculation."
Option Explicit

'@EntryPoint
'@Description "Calculates the date for a specific cell given a starting date."
Public Function CalculateDate(ByVal StartingDate As Date, ByVal Index As Long, ByVal Data As Range) As Date
Attribute CalculateDate.VB_Description = "Calculates the date for a specific cell given a starting date."
    'On first day, date is just starting date, after that it needs calculation.
    If Index = 1 Then
        CalculateDate = StartingDate
    Else
        'Date is dependent on remaining capacity, as with leftover capacity, new products can be started on the same day.
        Dim PreviousDate As Date
        PreviousDate = Data.Cells.Item(Index - 1, DateColumn)
        Dim PreviousProductAmount As Long
        PreviousProductAmount = Data.Cells.Item(Index - 1, AmountColumn)
        Dim RemainingCapacity As Long
        RemainingCapacity = Data.Cells.Item(Index - 1, RemainingCapacityColumn)
        If PreviousProductAmount <> 0 And RemainingCapacity > 0 Then
            CalculateDate = PreviousDate
        Else
            CalculateDate = DateAdd("d", 1, PreviousDate)
        End If
    End If
End Function
