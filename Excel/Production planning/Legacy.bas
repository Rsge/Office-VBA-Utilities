Attribute VB_Name = "Legacy"
Option Explicit

Private Function TextCurrentDate(CurrentDate As Date) As String
    TextCurrentDate = Format(CurrentDate, "dd.mm.yyyy")
End Function

Private Function TextToDate(CurrentDate As Date) As String
    TextToDate = "Am " + TextCurrentDate(CurrentDate) + ":  "
End Function

Public Function ProductionFinish(Capacity As Integer, Index As Integer, Data As Range) As String
    Dim CurrentDate As Date
    CurrentDate = CDate(Data.Cells(Index, 1))
    'If CurrentDate <> CDate("30.07.2021") Then Exit Function
    Dim ProductAmount As Integer
    ProductAmount = Data.Cells(Index, 3)
    Dim RemainingCapacity As Integer
    RemainingCapacity = Data.Cells(Index, 4)
    Dim Holiday As String
    Holiday = Data.Cells(Index, 5)
    Dim ItemNo As String
    ItemNo = Data.Cells(Index, 2)
    Dim i As Integer
    i = 1
    Dim Output As String
    Dim TempDate As Date
    
    If NoProduction(Holiday) Or RemainingCapacity > 0 Then
        ProductionFinish = ""
    ElseIf Index = 1 Then
        If RemainingCapacity = 0 Then
            ProductionFinish = TextToDate(CurrentDate) + ItemNo
        Else
            ProductionFinish = ""
        End If
    Else
        Dim LastRemainingCapacity As Integer
        LastRemainingCapacity = Data.Cells(Index - 1, 4)
        Dim FirstRemainingCapacity
        FirstRemainingCapacity = RemainingCapacity
        Dim PreviousDate As Date
        PreviousDate = CDate(Data.Cells(Index - 1, 1))
        If Capacity + LastRemainingCapacity >= 0 And ProductAmount <= Capacity Then
            If ProductAmount = 0 And RemainingCapacity < -Capacity Then
                ProductionFinish = ""
            Else
                If RemainingCapacity >= 0 Then Output = ItemNo
                If ProductAmount <> Capacity Then
                    Do
                        ItemNo = Data.Cells(Index - i, 2)
                        RemainingCapacity = Data(Index - i, 4)
                        If ItemNo <> "" And RemainingCapacity <> 0 Then Output = ItemNo + ", " + Output
                        TempDate = CDate(Data.Cells(Index - i, 1))
                        i = i + 1
                        PreviousDate = CDate(Data.Cells(Index - i, 1))
                    Loop Until PreviousDate <> TempDate Or RemainingCapacity <= 0
                    Holiday = Data.Cells(Index - i, 5)
                    If RemainingCapacity <> 0 Or (NoProduction(Holiday) And FirstRemainingCapacity >= 0) Then
                        Do
                            ItemNo = Data.Cells(Index - i, 2)
                            i = i + 1
                        Loop Until ItemNo <> ""
                    End If
                End If
                RemainingCapacity = Data.Cells(Index - i + 1, 4)
                If RemainingCapacity < 0 Then Output = ItemNo + ", " + Output
                If Right(Output, 2) = ", " Then Output = Left(Output, Len(Output) - 2)
                If Not Output = "" Then
                    ProductionFinish = TextToDate(CurrentDate) + Output
                Else
                    ProductionFinish = ""
                End If
            End If
        ElseIf Capacity + LastRemainingCapacity >= 0 And Abs(ProductAmount - Capacity) < Abs(RemainingCapacity) Then
            Do
                ItemNo = Data.Cells(Index - i, 2)
                i = i + 1
            Loop Until ItemNo <> ""
            RemainingCapacity = Data.Cells(Index - i + 1, 4)
            Output = ItemNo + ", " + Output
            If Right(Output, 2) = ", " Then Output = Left(Output, Len(Output) - 2)
            ProductionFinish = TextToDate(CurrentDate) + Output
        End If
    End If
End Function


Public Function ShowDueItems(Index As Integer, Jobs As Range, Data As Range)
    Dim JobData As Object
    Set JobData = CreateObject("Scripting.Dictionary")
    Dim ItemData As Object
    Set ItemData = CreateObject("Scripting.Dictionary")
    Dim CurrentRow As Range
    For Each CurrentRow In Jobs.Rows
        Dim JobNo As String
        JobNo = CurrentRow.Cells(1, 1)
        Dim ItemNo As String
        ItemNo = CurrentRow.Cells(1, 2)
        Dim Due As Date
        Due = CurrentRow.Cells(1, 3)
        JobData(Due) = JobData(Due) + ItemNo + ", "
    Next
    
    Dim CurrentDate As Date
    CurrentDate = Data.Cells(Index, 1)
    If JobData.Exists(CurrentDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells(Index + 1, 1)
        If NextDay <> CurrentDate Then
            ShowDueItems = Left(JobData(CurrentDate), Len(JobData(CurrentDate)) - 2)
        Else
            ShowDueItems = ""
        End If
    Else
        ShowDueItems = ""
    End If
End Function

