Attribute VB_Name = "DateCalc"
Option Explicit

Public Function CalculateDate(StartingDate As Date, Index As Integer, Data As Range) As Date
    If Index = 1 Then
        CalculateDate = StartingDate
    Else
        Dim PreviousDate As Date
        PreviousDate = Data.Cells(Index - 1, 1)
        Dim PreviousProductAmount As Integer
        PreviousProductAmount = Data.Cells(Index - 1, 4)
        Dim RemainingCapacity As Integer
        RemainingCapacity = Data.Cells(Index - 1, 7)
        If PreviousProductAmount <> 0 And RemainingCapacity > 0 Then
            CalculateDate = PreviousDate
        Else
            CalculateDate = DateAdd("d", 1, PreviousDate)
        End If
    End If
End Function

