Attribute VB_Name = "DateCalc"
Attribute VB_Description = "Date calculation."
'@Folder "Production planning"
'@ModuleDescription "Date calculation."
Option Explicit

'@EntryPoint
'@Description "Calculates the date for a specific cell given a starting date."
Public Function CalculateDate(ByVal startingDate As Date, ByVal index As Long, ByVal data As Range) As Date
Attribute CalculateDate.VB_Description = "Calculates the date for a specific cell given a starting date."
    'On first day, date is just starting date, after that it needs calculation.
    If index = 1 Then
        CalculateDate = startingDate
    Else
        'Date is dependent on remaining capacity, as with leftover capacity, new products can be started on the same day.
        Dim previousDate As Date
        previousDate = data.Cells.Item(index - 1, DateColumn)
        Dim previousProductAmount As Long
        previousProductAmount = data.Cells.Item(index - 1, AmountColumn)
        Dim remainingCapacity As Long
        remainingCapacity = data.Cells.Item(index - 1, RemainingCapacityColumn)
        If previousProductAmount <> 0 And remainingCapacity > 0 Then
            CalculateDate = previousDate
        Else
            CalculateDate = DateAdd("d", 1, previousDate)
        End If
    End If
End Function
