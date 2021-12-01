Attribute VB_Name = "DoneJobs"
Option Explicit

Private DueJobs As Object
Private DoneJobs As Object

Public Function ShowDueJobs(Index As Integer, Jobs As Range, Data As Range) As String
    If Index = 1 Then
        Set DueJobs = CreateObject("Scripting.Dictionary")
        Dim JobRow As Range
        For Each JobRow In Jobs.Rows
            If JobRow.Cells(1, 1) <> "" Then
                Dim Job As String
                Job = JobRow.Cells(1, 1)
                Dim Due As Date
                Due = JobRow.Cells(1, 2)
                DueJobs(Due) = DueJobs(Due) & Job & ", "
            End If
        Next
    End If
    Dim DueDate As Date
    DueDate = Data.Cells(Index, 1)
    If DueJobs.Exists(DueDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells(Index + 1, 1)
        If NextDay <> DueDate Then
            ShowDueJobs = Left(DueJobs(DueDate), Len(DueJobs(DueDate)) - 2)
        Else
            ShowDueJobs = ""
        End If
    Else
        ShowDueJobs = ""
    End If
End Function


Public Function EarliestJobCompletion(Capacity As Integer, Index As Integer, Data As Range) As String
    If Index = 1 Then
        'Set ProduktAmount = CreateObject("Scripting.Dictionary")
        Set DoneJobs = CreateObject("Scripting.Dictionary")
    End If
    Dim DueDate As Date
    DueDate = Data.Cells(Index, 1)
    'If DueDate < CDate("19.05.2021") Then Exit Function
    Dim Job As String
    Job = Data.Cells(Index, 2)
    'Dim ProductAmount As Integer
    'ProductAmount = Data.Cells(Index, 4)
    Dim i As Integer
    i = 1
    
    If Job <> "" Then
        'ProduktAmount(Job) = ProduktAmount(Job) + ProductAmount
        Do While (Index + i) <= Data.Rows.Count And Data.Cells(Index + i, 2) = ""
'            If Data.Cells(Index + i, 2) = Job Then
'                IsLastEntry = False
'                Exit Do
'            End If
            i = i + 1
        Loop
        If Data.Cells(Index + i, 2) <> Job Then
            Dim RemainingCapacity As Long
            RemainingCapacity = Data.Cells(Index, 7)
            If RemainingCapacity >= 0 Then
                If Not InStr(DoneJobs(DueDate), Job) Then
                    DoneJobs(DueDate) = DoneJobs(DueDate) & Job & ", "
                End If
            Else
                i = 0
                Dim RemainingProduction As Long
                RemainingProduction = Abs(RemainingCapacity)
                Dim Ende As Boolean
                Dim Done As Date
                Do
                    i = i + 1
                    Done = DateAdd("d", i, DueDate)
                    If Not NoProduction(Done) Then
                        RemainingProduction = RemainingProduction - Capacity
                    End If
                    Ende = (Index + i) >= Data.Rows.Count
                Loop While RemainingProduction > 0 And Not Ende
                
                If RemainingProduction > 0 Or Index + i > Data.Rows.Count Then
                    Done = Data.Cells(Data.Rows.Count, 1)
                    If InStr(DoneJobs(Done), "i.Z.") Then
                        DoneJobs(Done) = DoneJobs(Done) & Job & ", "
                    Else
                        DoneJobs(Done) = DoneJobs(Done) & "i.Z.: " & Job & ", "
                    End If
                Else
                    DoneJobs(Done) = DoneJobs(Done) & Job & ", "
                End If
            End If
            'ProduktAmount(Job) = 0
        End If
    End If
    
    If DoneJobs.Exists(DueDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells(Index + 1, 1)
        If NextDay <> DueDate Then
            EarliestJobCompletion = Left(DoneJobs(DueDate), Len(DoneJobs(DueDate)) - 2)
        Else
            EarliestJobCompletion = ""
        End If
    Else
        EarliestJobCompletion = ""
    End If
End Function

