Attribute VB_Name = "DateCalculation"
Attribute VB_Description = "Date calculation."
'@Folder("ProductionPlanning.Dates")
'@ModuleDescription("Date calculation.")
Option Explicit

'@EntryPoint
'@Description("Calculates the date for a specific cell given a starting date.")
Public Function CalulateDate(ByVal startingDate As Date, ByVal index As Long, ByVal data As Range) As Date
Attribute CalulateDate.VB_Description = "Calculates the date for a specific cell given a starting date."
    ' On first day, date is just starting date, after that it needs calculation.
    If index = 1 Then
        CalulateDate = startingDate
    Else
        ' Date is dependent on remaining capacity, as with leftover capacity, new products can be started on the same day.
        Dim previousDate As Date
        previousDate = GetCellValue(data, index - 1, DateColumn)
        Dim previousProductAmount As Long
        previousProductAmount = GetCellValue(data, index - 1, AmountColumn)
        Dim remainingCapacity As Long
        remainingCapacity = GetCellValue(data, index - 1, RemainingCapacityColumn)
        If previousProductAmount <> 0 And remainingCapacity > 0 Then
            CalulateDate = previousDate
        Else
            CalulateDate = DateAdd("d", 1, previousDate)
        End If
    End If
End Function
