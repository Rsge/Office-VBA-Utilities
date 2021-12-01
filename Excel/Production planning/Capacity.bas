Attribute VB_Name = "Capacity"
Option Explicit

Public Function CalculateCapacity(Capacity As Integer, Index As Integer, Data As Range) As Long
    Dim RowDate As Date
    RowDate = Data.Cells(Index, 1)
    If KeineProduktion(RowDate) Then
        CalculateCapacity = 0
    Else
        Dim ProductAmount As Integer
        ProductAmount = Data.Cells(Index, 4)
        Dim Slowdown As Integer
        Slowdown = Data.Cells(Index, 9)
        If Index = 1 Then
            CalculateCapacity = Capacity - ProductAmount - Slowdown
        Else
            Dim CurrentDate As Date
            CurrentDate = Data.Cells(Index, 1)
            Dim PreviousDate As Date
            Dim PreviousProductAmount As Integer
            Dim RemainingCapacity As Integer
            Dim PreviousRemainingCapacity As Integer
            PreviousRemainingCapacity = Capacity
            Dim i As Integer
            i = 1
            Do
                PreviousDate = Data.Cells(Index - i, 1)
                PreviousProductAmount = Data.Cells(Index - i, 4)
                RemainingCapacity = Data.Cells(Index - i, 7)
                i = i + 1
                If Index - i > 1 Then PreviousRemainingCapacity = Data.Cells(Index - i, 7)
            Loop While KeineProduktion(PreviousDate) And Index - i > 1
            If RemainingCapacity = Capacity Or (PreviousProductAmount = 0 And RemainingCapacity >= 0) Then
                CalculateCapacity = Capacity - ProductAmount - Slowdown
            ElseIf CurrentDate = PreviousDate Then
                CalculateCapacity = RemainingCapacity - ProductAmount - Slowdown
            Else
                CalculateCapacity = Capacity + RemainingCapacity - ProductAmount - Slowdown
            End If
        End If
    End If
End Function

