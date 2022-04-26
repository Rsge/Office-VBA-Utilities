Attribute VB_Name = "Jobs"
Attribute VB_Description = "Job calculations."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Job calculations."
Option Explicit

'String constant
'@VariableDescription "Info string added to job on last day to indicate it'll finish in the future, beyond the current table's scopes."
Private Const m_futureInfo As String = "Future:"
Attribute m_futureInfo.VB_VarDescription = "Info string added to job on last day to indicate it'll finish in the future, beyond the current table's scopes."
'Column constants
'@VariableDescription "Jobs' numeric identifiers' column."
Private Const m_jobNumColumn As Long = 1
Attribute m_jobNumColumn.VB_VarDescription = "Jobs' numeric identifiers' column."
'@VariableDescription "Jobs' dates' column."
Private Const m_jobDueDateColumn As Long = 2
Attribute m_jobDueDateColumn.VB_VarDescription = "Jobs' dates' column."

'Variables
'@VariableDescription "Dictionary of jobs due at their key's date."
'@Ignore MoveFieldCloserToUsage
Private m_dueJobs As Object
Attribute m_dueJobs.VB_VarDescription = "Dictionary of jobs due at their key's date."
'@VariableDescription "Dictionary of jobs probably done at their key's date."
'@Ignore MoveFieldCloserToUsage
Private m_doneJobs As Object
Attribute m_doneJobs.VB_VarDescription = "Dictionary of jobs probably done at their key's date."

'@EntryPoint
'@Description "Shows due jobs in row of their deadline date."
Public Function ShowDueJobs(ByVal index As Long, ByVal jobs As Range, ByVal data As Range) As String
Attribute ShowDueJobs.VB_Description = "Shows due jobs in row of their deadline date."
    'Getting due jobs' info only on processing at first index
    If index = 1 Then
        Set m_dueJobs = CreateObject("Scripting.Dictionary")
        Dim jobRow As Range
        For Each jobRow In jobs.Rows
            If jobRow.Cells.Item(1, m_jobNumColumn) <> vbNullString Then
                Dim job As String
                job = jobRow.Cells.Item(1, m_jobNumColumn)
                Dim due As Date
                due = jobRow.Cells.Item(1, m_jobDueDateColumn)
                m_dueJobs(due) = m_dueJobs(due) & job & Comma
            End If
        Next
    End If
    'Showing due job(s) on it's/their date
    Dim dueDate As Date
    dueDate = data.Cells.Item(index, DateColumn)
    If m_dueJobs.Exists(dueDate) Then
        Dim nextDay As Date
        nextDay = data.Cells.Item(index + 1, DateColumn)
        If nextDay <> dueDate Then
            ShowDueJobs = Left$(m_dueJobs(dueDate), Len(m_dueJobs(dueDate)) - 2)
        Else
            ShowDueJobs = vbNullString
        End If
    Else
        ShowDueJobs = vbNullString
    End If
End Function


'@EntryPoint
'@Description "Shows jobs in row of their respective earliest completion date according to current inputs."
Public Function EarliestJobCompletion(ByVal baseCapacity As Long, ByVal index As Long, ByVal data As Range) As String
Attribute EarliestJobCompletion.VB_Description = "Shows jobs in row of their respective earliest completion date according to current inputs."
    'Variables
    If index = 1 Then
        Set m_doneJobs = CreateObject("Scripting.Dictionary")
    End If
    Dim dueDate As Date
    dueDate = data.Cells.Item(index, DateColumn)
    Dim job As String
    job = data.Cells.Item(index, JobColumn)
    Dim i As Long
    i = 1
    
    'Calculating when jobs will be done
    If job <> vbNullString Then
        Do While (index + i) <= data.Rows.Count And LenB(data.Cells.Item(index + i, JobColumn)) = 0
            i = i + 1
        Loop
        If data.Cells.Item(index + i, 2) <> job Then
            Dim remainingCapacity As Long
            remainingCapacity = data.Cells.Item(index, RemainingCapacityColumn)
            'If there is remaining capacity, the job is done and can be added, otherwise further calculation is needed
            If remainingCapacity >= 0 Then
                If Not InStr(m_doneJobs(dueDate), job) Then
                    m_doneJobs(dueDate) = m_doneJobs(dueDate) & job & Comma
                End If
            Else
                i = 0
                Dim remainingProduction As Long
                remainingProduction = Abs(remainingCapacity)
                Dim ending As Boolean
                Dim doneDate As Date
                'Finding if a job can be done in the current timeframe of the data table or if it's later than that, using base capacity.
                Do
                    i = i + 1
                    doneDate = DateAdd("d", i, dueDate)
                    If Not NoProduction(doneDate) Then
                        remainingProduction = remainingProduction - baseCapacity
                    End If
                    ending = (index + i) >= data.Rows.Count
                Loop While remainingProduction > 0 And Not ending
                
                'Adding job at it's approximate date to the dictionary
                If remainingProduction > 0 Or index + i > data.Rows.Count Then
                    doneDate = data.Cells.Item(data.Rows.Count, 1)
                    If InStr(m_doneJobs(doneDate), m_futureInfo) Then
                        m_doneJobs(doneDate) = m_doneJobs(doneDate) & job & Comma
                    Else
                        m_doneJobs(doneDate) = m_doneJobs(doneDate) & m_futureInfo & Colon & Space(1) & job & Comma
                    End If
                Else
                    m_doneJobs(doneDate) = m_doneJobs(doneDate) & job & Comma
                End If
            End If
        End If
    End If
    
    'Showing jobs on their potential completion dates or on end of table
    If m_doneJobs.Exists(dueDate) Then
        Dim nextDay As Date
        nextDay = data.Cells.Item(index + 1, DateColumn)
        If nextDay <> dueDate Then
            EarliestJobCompletion = Left$(m_doneJobs(dueDate), Len(m_doneJobs(dueDate)) - 2)
        Else
            EarliestJobCompletion = vbNullString
        End If
    Else
        EarliestJobCompletion = vbNullString
    End If
End Function
