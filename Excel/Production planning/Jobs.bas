Attribute VB_Name = "Jobs"
Attribute VB_Description = "Module for methods related to jobs."
'@IgnoreModule InvalidAnnotation, IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Module for methods related to jobs."
Option Explicit

'String constant
'@VariableDescription "Info string added to job on last day to indicate it'll finish in the future, beyond the current table's scopes."
Private Const FutureInfo As String = "Future:"
'Column constants
'@VariableDescription "Jobs' numeric identifiers' column."
Private Const JobNumColumn As Long = 1
'@VariableDescription "Jobs' dates' column."
Private Const JobDueDateColumn As Long = 2

'Variables
'@VariableDescription "Dictionary of jobs due at their key's date."
'@Ignore MoveFieldCloserToUsage
Private DueJobs As Object
Attribute DueJobs.VB_VarDescription = "Dictionary of jobs due at their key's date."
'@VariableDescription "Dictionary of jobs probably done at their key's date."
'@Ignore MoveFieldCloserToUsage
Private DoneJobs As Object
Attribute DoneJobs.VB_VarDescription = "Dictionary of jobs probably done at their key's date."

'@EntryPoint
'@Description "Shows due jobs in row of their deadline date."
Public Function ShowDueJobs(ByVal Index As Long, ByVal Jobs As Range, ByVal Data As Range) As String
Attribute ShowDueJobs.VB_Description = "Shows due jobs in row of their deadline date."
    'Get due jobs' info only on processing at first index
    If Index = 1 Then
        Set DueJobs = CreateObject("Scripting.Dictionary")
        Dim JobRow As Range
        For Each JobRow In Jobs.Rows
            If JobRow.Cells.Item(1, JobNumColumn) <> vbNullString Then
                Dim Job As String
                Job = JobRow.Cells.Item(1, JobNumColumn)
                Dim Due As Date
                Due = JobRow.Cells.Item(1, JobDueDateColumn)
                DueJobs(Due) = DueJobs(Due) & Job & Comma
            End If
        Next
    End If
    'Show due job(s) on it's/their date
    Dim DueDate As Date
    DueDate = Data.Cells.Item(Index, DateColumn)
    If DueJobs.Exists(DueDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells.Item(Index + 1, DateColumn)
        If NextDay <> DueDate Then
            ShowDueJobs = Left$(DueJobs(DueDate), Len(DueJobs(DueDate)) - 2)
        Else
            ShowDueJobs = vbNullString
        End If
    Else
        ShowDueJobs = vbNullString
    End If
End Function


'@EntryPoint
'@Description "Shows jobs in row of their respective earliest completion date according to current inputs."
Public Function EarliestJobCompletion(ByVal BaseCapacity As Long, ByVal Index As Long, ByVal Data As Range) As String
Attribute EarliestJobCompletion.VB_Description = "Shows jobs in row of their respective earliest completion date according to current inputs."
    'Variables
    If Index = 1 Then
        Set DoneJobs = CreateObject("Scripting.Dictionary")
    End If
    Dim DueDate As Date
    DueDate = Data.Cells.Item(Index, DateColumn)
    Dim Job As String
    Job = Data.Cells.Item(Index, JobColumn)
    Dim i As Long
    i = 1
    
    'Calculate, when jobs will be done
    If Job <> vbNullString Then
        Do While (Index + i) <= Data.Rows.Count And LenB(Data.Cells.Item(Index + i, JobColumn)) = 0
            i = i + 1
        Loop
        If Data.Cells.Item(Index + i, 2) <> Job Then
            Dim RemainingCapacity As Long
            RemainingCapacity = Data.Cells.Item(Index, RemainingCapacityColumn)
            'If there is remaining capacity, the job is done and can be added, otherwise further calculation is needed
            If RemainingCapacity >= 0 Then
                If Not InStr(DoneJobs(DueDate), Job) Then
                    DoneJobs(DueDate) = DoneJobs(DueDate) & Job & Comma
                End If
            Else
                i = 0
                Dim RemainingProduction As Long
                RemainingProduction = Abs(RemainingCapacity)
                Dim Ending As Boolean
                Dim DoneDate As Date
                'Find if a job can be done in the current timeframe of the data table or if it's later than that, using base capacity.
                Do
                    i = i + 1
                    DoneDate = DateAdd("d", i, DueDate)
                    If Not NoProduction(DoneDate) Then
                        RemainingProduction = RemainingProduction - BaseCapacity
                    End If
                    Ending = (Index + i) >= Data.Rows.Count
                Loop While RemainingProduction > 0 And Not Ending
                
                'Add job at it's approximate date to the dictionary
                If RemainingProduction > 0 Or Index + i > Data.Rows.Count Then
                    DoneDate = Data.Cells.Item(Data.Rows.Count, 1)
                    If InStr(DoneJobs(DoneDate), FutureInfo) Then
                        DoneJobs(DoneDate) = DoneJobs(DoneDate) & Job & Comma
                    Else
                        DoneJobs(DoneDate) = DoneJobs(DoneDate) & FutureInfo & Colon & Space(1) & Job & Comma
                    End If
                Else
                    DoneJobs(DoneDate) = DoneJobs(DoneDate) & Job & Comma
                End If
            End If
        End If
    End If
    
    'Show jobs on their potential completion dates or on end of table
    If DoneJobs.Exists(DueDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells.Item(Index + 1, DateColumn)
        If NextDay <> DueDate Then
            EarliestJobCompletion = Left$(DoneJobs(DueDate), Len(DoneJobs(DueDate)) - 2)
        Else
            EarliestJobCompletion = vbNullString
        End If
    Else
        EarliestJobCompletion = vbNullString
    End If
End Function
