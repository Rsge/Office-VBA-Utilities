Attribute VB_Name = "Jobs"
Attribute VB_Description = "Job calculations."
'@Folder("ProductionPlanning.AdditionalInfo")
'@ModuleDescription("Job calculations.")
Option Explicit

' Column constants
'@VariableDescription("Jobs' numeric identifiers' column.")
Private Const m_jobNumColumn As Long = 1
Attribute m_jobNumColumn.VB_VarDescription = "Jobs' numeric identifiers' column."
'@VariableDescription("Jobs' dates' column.")
Private Const m_jobDueDateColumn As Long = 2
Attribute m_jobDueDateColumn.VB_VarDescription = "Jobs' dates' column."

' Variables
'@VariableDescription("Dictionary of jobs due at their key's date.")
'@Ignore MoveFieldCloserToUsage
Private m_dueJobs As Object
Attribute m_dueJobs.VB_VarDescription = "Dictionary of jobs due at their key's date."
'@VariableDescription("Dictionary of jobs probably done at their key's date.")
'@Ignore MoveFieldCloserToUsage
Private m_doneJobs As Object
Attribute m_doneJobs.VB_VarDescription = "Dictionary of jobs probably done at their key's date."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Shows due jobs in row of their deadline date.")
Public Function ShowDueJobs(ByVal index As Long, ByVal jobs As Range, ByVal data As Range) As String
Attribute ShowDueJobs.VB_Description = "Shows due jobs in row of their deadline date."
    ' Get due jobs' info only on processing at first index.
    If index = 1 Then
        Set m_dueJobs = CreateObject("Scripting.Dictionary")
        Dim jobRow As Range
        For Each jobRow In jobs.Rows
            If GetCellValue(jobRow, 1, m_jobNumColumn) <> vbNullString Then
                Dim job As String
                job = GetCellValue(jobRow, 1, m_jobNumColumn)
                Dim due As Date
                due = GetCellValue(jobRow, 1, m_jobDueDateColumn)
                m_dueJobs.Item(due) = m_dueJobs.Item(due) & job & Comma
            End If
        Next
    End If
    ' Show due job(s) on its/their date.
    Dim dueDate As Date
    dueDate = CDate(GetCellValue(data, index, DateColumn))
    If m_dueJobs.Exists(dueDate) Then
        Dim nextDay As Date
        nextDay = GetCellValue(data, index + 1, DateColumn)
        If nextDay <> dueDate Then
            ShowDueJobs = Left$(m_dueJobs.Item(dueDate), Len(m_dueJobs.Item(dueDate)) - 2)
        Else
            ShowDueJobs = vbNullString
        End If
    Else
        ShowDueJobs = vbNullString
    End If
End Function


'@EntryPoint
'@Description("Shows jobs in row of their respective earliest completion date according to current inputs.")
Public Function EarliestJobCompletion(ByVal baseCapacity As Long, ByVal index As Long, ByVal data As Range) As String
Attribute EarliestJobCompletion.VB_Description = "Shows jobs in row of their respective earliest completion date according to current inputs."
    ' Variables
    If index = 1 Then
        Set m_doneJobs = CreateObject("Scripting.Dictionary")
    End If
    Dim dueDate As Date
    dueDate = GetCellValue(data, index, DateColumn)
    Dim job As String
    job = GetCellValue(data, index, JobColumn)
    Dim i As Long
    i = 1
    
    ' Calculate when jobs will be done.
    If job <> vbNullString Then
        Do While (index + i) <= data.Rows.Count And IsEmpty(GetCellValue(data, index + i, JobColumn))
            i = i + 1
        Loop
        If GetCellValue(data, index + i, 2) <> job Then
            Dim remainingCapacity As Long
            remainingCapacity = GetCellValue(data, index, RemainingCapacityColumn)
            ' If there is remaining capacity, the job is done and can be added, otherwise further calculation is needed.
            If remainingCapacity >= 0 Then
                If Not InStr(m_doneJobs.Item(dueDate), job) Then
                    m_doneJobs.Item(dueDate) = m_doneJobs.Item(dueDate) & job & Comma
                End If
            Else
                i = 0
                Dim remainingProduction As Long
                remainingProduction = Abs(remainingCapacity)
                Dim ending As Boolean
                Dim doneDate As Date
                ' Find if a job can be done in the current timeframe of the data table or if it's later than that, use base capacity.
                Do
                    i = i + 1
                    doneDate = DateAdd("d", i, dueDate)
                    If Not NoProduction(doneDate) Then
                        remainingProduction = remainingProduction - baseCapacity
                    End If
                    ending = (index + i) >= data.Rows.Count
                Loop While remainingProduction > 0 And Not ending
                
                ' Add job at it's approximate date to the dictionary.
                If remainingProduction > 0 Or index + i > data.Rows.Count Then
                    doneDate = GetCellValue(data, data.Rows.Count, 1)
                    If InStr(m_doneJobs.Item(doneDate), FutureInfo) Then
                        m_doneJobs.Item(doneDate) = m_doneJobs.Item(doneDate) & job & Comma
                    Else
                        m_doneJobs.Item(doneDate) = m_doneJobs.Item(doneDate) & FutureInfo & Colon & Space(1) & job & Comma
                    End If
                Else
                    m_doneJobs.Item(doneDate) = m_doneJobs.Item(doneDate) & job & Comma
                End If
            End If
        End If
    End If
    
    ' Show jobs on their potential completion dates or on end of table.
    If m_doneJobs.Exists(dueDate) Then
        Dim nextDay As Date
        nextDay = GetCellValue(data, index + 1, DateColumn)
        If nextDay <> dueDate Then
            EarliestJobCompletion = Left$(m_doneJobs.Item(dueDate), Len(m_doneJobs.Item(dueDate)) - 2)
        Else
            EarliestJobCompletion = vbNullString
        End If
    Else
        EarliestJobCompletion = vbNullString
    End If
End Function
