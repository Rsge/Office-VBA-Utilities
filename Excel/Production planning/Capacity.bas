Attribute VB_Name = "Capacity"
Attribute VB_Description = "Production capacity calculations."
'@Folder "Production planning"
'@ModuleDescription "Production capacity calculations."
Option Explicit

'@EntryPoint
'@Description "Calculates the capacity for a specified date and slowdown and the amount to produce."
Public Function CalculateCapacity(ByVal baseCapacity As Long, ByVal index As Long, ByVal data As Range) As Long
Attribute CalculateCapacity.VB_Description = "Calculates the capacity for a specified date and slowdown and the amount to produce."
    Dim currentDate As Date
    currentDate = data.Cells.Item(index, DateColumn)
    'On no-production-dates, capacity is zero, otherwise it needs to be calculated
    If NoProduction(currentDate) Then
        CalculateCapacity = 0
    Else
        'Getting amount and slowdown
        Dim productAmount As Long
        productAmount = data.Cells.Item(index, AmountColumn)
        Dim slowdown As Long
        slowdown = data.Cells.Item(index, SlowdownsColumn)
        'On first index, only base values matter, otherwise previous data has to be accounted for
        If index = 1 Then
            CalculateCapacity = baseCapacity - productAmount - slowdown
        Else
            Dim previousDate As Date
            Dim previousProductAmount As Long
            Dim remainingCapacity As Long
            Dim i As Long
            i = 1
            'Skipping no-production-dates
            Do
                previousDate = data.Cells.Item(index - i, DateColumn)
                previousProductAmount = data.Cells.Item(index - i, AmountColumn)
                remainingCapacity = data.Cells.Item(index - i, RemainingCapacityColumn)
                i = i + 1
            Loop While NoProduction(previousDate) And index - i > 1
            'Calculating remaining capacity for different scenarios
            If remainingCapacity = baseCapacity Or (previousProductAmount = 0 And remainingCapacity >= 0) Then
                CalculateCapacity = baseCapacity - productAmount - slowdown
            ElseIf currentDate = previousDate Then
                CalculateCapacity = remainingCapacity - productAmount - slowdown
            Else
                CalculateCapacity = baseCapacity + remainingCapacity - productAmount - slowdown
            End If
        End If
    End If
End Function
