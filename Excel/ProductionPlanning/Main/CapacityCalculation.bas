Attribute VB_Name = "CapacityCalculation"
Attribute VB_Description = "Production capacity calculations."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("Production capacity calculations.")
Option Explicit

'@EntryPoint
'@Description("Calculates the capacity for a specified date and slowdown and the amount to produce.")
Public Function CalculateCapacity(ByVal baseCapacity As Long, ByVal index As Long, ByVal data As Range) As Long
Attribute CalculateCapacity.VB_Description = "Calculates the capacity for a specified date and slowdown and the amount to produce."
    Dim currentDate As Date
    currentDate = CDate(GetCellValue(data, index, DateColumn))
    ' On no-production-dates, capacity is zero, otherwise it needs to be calculated.
    If NoProduction(currentDate) Then
        CalculateCapacity = 0
    Else
        ' Get amount and slowdown.
        Dim productAmount As Long
        productAmount = GetCellValue(data, index, AmountColumn)
        Dim Slowdown As Long
        Slowdown = GetCellValue(data, index, SlowdownsColumn)
        ' On first index, only base values matter, otherwise previous data has to be accounted for.
        If index = 1 Then
            CalculateCapacity = baseCapacity - productAmount - Slowdown
        Else
            Dim previousDate As Date
            Dim previousProductAmount As Long
            Dim remainingCapacity As Long
            Dim i As Long
            i = 1
            ' Skip no-production-dates.
            Do
                previousDate = GetCellValue(data, index - i, DateColumn)
                previousProductAmount = GetCellValue(data, index - i, AmountColumn)
                remainingCapacity = GetCellValue(data, index - i, RemainingCapacityColumn)
                i = i + 1
            Loop While NoProduction(previousDate) And index - i > 1
            ' Calculate remaining capacity for different scenarios.
            If remainingCapacity = baseCapacity Or (previousProductAmount = 0 And remainingCapacity >= 0) Then
                CalculateCapacity = baseCapacity - productAmount - Slowdown
            ElseIf currentDate = previousDate Then
                CalculateCapacity = remainingCapacity - productAmount - Slowdown
            Else
                CalculateCapacity = baseCapacity + remainingCapacity - productAmount - Slowdown
            End If
        End If
    End If
End Function
