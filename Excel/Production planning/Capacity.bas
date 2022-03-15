Attribute VB_Name = "Capacity"
Attribute VB_Description = "Module for methods related to production capacity."
'@Folder "Production planning"
'@ModuleDescription "Module for methods related to production capacity."
Option Explicit

'@EntryPoint
'@Description "Calculates the capacity for a specified date and slowdown and the amount to produce."
Public Function CalculateCapacity(ByVal BaseCapacity As Long, ByVal Index As Long, ByVal Data As Range) As Long
Attribute CalculateCapacity.VB_Description = "Calculates the capacity for a specified date and slowdown and the amount to produce."
    Dim CurrentDate As Date
    CurrentDate = Data.Cells.Item(Index, DateColumn)
    'On no-production-dates, capacity is zero, otherwise it needs to be calculated
    If NoProduction(CurrentDate) Then
        CalculateCapacity = 0
    Else
        'Get amount and slowdown
        Dim ProductAmount As Long
        ProductAmount = Data.Cells.Item(Index, AmountColumn)
        Dim Slowdown As Long
        Slowdown = Data.Cells.Item(Index, SlowdownsColumn)
        'On first index, only base values matter, otherwise previous data has to be accounted for
        If Index = 1 Then
            CalculateCapacity = BaseCapacity - ProductAmount - Slowdown
        Else
            Dim PreviousDate As Date
            Dim PreviousProductAmount As Long
            Dim RemainingCapacity As Long
            Dim i As Long
            i = 1
            'Skip no-production-dates
            Do
                PreviousDate = Data.Cells.Item(Index - i, DateColumn)
                PreviousProductAmount = Data.Cells.Item(Index - i, AmountColumn)
                RemainingCapacity = Data.Cells.Item(Index - i, RemainingCapacityColumn)
                i = i + 1
            Loop While NoProduction(PreviousDate) And Index - i > 1
            'Calculate remaining capacity for different scenarios
            If RemainingCapacity = BaseCapacity Or (PreviousProductAmount = 0 And RemainingCapacity >= 0) Then
                CalculateCapacity = BaseCapacity - ProductAmount - Slowdown
            ElseIf CurrentDate = PreviousDate Then
                CalculateCapacity = RemainingCapacity - ProductAmount - Slowdown
            Else
                CalculateCapacity = BaseCapacity + RemainingCapacity - ProductAmount - Slowdown
            End If
        End If
    End If
End Function
