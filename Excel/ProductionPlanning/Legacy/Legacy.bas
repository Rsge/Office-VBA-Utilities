Attribute VB_Name = "Legacy"
Attribute VB_Description = "Legacy code saving."
'@Folder("ProductionPlanning.Legacy")
'@ModuleDescription("Legacy code saving.")
Option Explicit

' String constants
'@VariableDescription("Prefix for a date string to symbolize something happens on this day.)
Private Const m_datePrefix As String = "On "

' ————————————————————————————————————————————————————— '


' Date to string functions
'@Description("Converts a date to a string in default european date format.)
Private Function TextCurrentDate(ByVal currentDate As Date) As String
    TextCurrentDate = Format$(currentDate, "dd.mm.yyyy")
End Function

'@Description("Returns an info text with the given date.")
Private Function TextToDate(ByVal currentDate As Date) As String
Attribute TextToDate.VB_Description = "Returns an info text with the given date."
    TextToDate = m_datePrefix & TextCurrentDate(currentDate) & Colon & Space(1)
End Function

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Calculates the point at which the production of an item would finish")
Private Function ProductionFinish(ByVal Capacity As Long, ByVal index As Long, ByVal data As Range) As String
Attribute ProductionFinish.VB_Description = "Calculates the point at which the production of an item would finish"
    ' Variables
    Dim currentDate As Date
    currentDate = CDate(data.Cells.Item(index, DateColumn))
    Dim productAmount As Long
    productAmount = data.Cells.Item(index, AmountColumn)
    Dim remainingCapacity As Long
    remainingCapacity = data.Cells.Item(index, RemainingCapacityColumn)
    Dim holiday As String
    holiday = data.Cells.Item(index, HolidaysColumn)
    Dim itemNum As String
    itemNum = data.Cells.Item(index, ItemColumn)
    Dim i As Long
    i = 1
    Dim output As String
    Dim tempDate As Date
    
    ' If no production is done or there is remaining capacity, item isn't done,
    ' else if it's the first entry, item is only done when there's no remaining capacity,
    ' else it has to be further calculated.
    If LenB(holiday) = 0 Or remainingCapacity > 0 Then
        ProductionFinish = vbNullString
    ElseIf index = 1 Then
        If remainingCapacity = 0 Then
            ProductionFinish = TextToDate(currentDate) & itemNum
        Else
            ProductionFinish = vbNullString
        End If
    Else
        Dim lastRemainingCapacity As Long
        lastRemainingCapacity = data.Cells.Item(index - 1, RemainingCapacityColumn)
        Dim firstRemainingCapacity As Long
        firstRemainingCapacity = remainingCapacity
        Dim previousDate As Date
        previousDate = CDate(data.Cells.Item(index - 1, DateColumn))
        'If there's enough capacity to finish the current item, it can be added as finished
        If Capacity + lastRemainingCapacity >= 0 And productAmount <= Capacity Then
            If productAmount = 0 And remainingCapacity < -Capacity Then
                ProductionFinish = vbNullString
            Else
                ' If there's remaining capacity, the current item is finished.
                If remainingCapacity >= 0 Then output = itemNum
                ' And if there's enough capacity, older items will be finished now, too.
                If productAmount <> Capacity Then
                    Do
                        itemNum = data.Cells.Item(index - i, ItemColumn)
                        remainingCapacity = data.Item(index - i, RemainingCapacityColumn)
                        If itemNum <> vbNullString And remainingCapacity <> 0 Then output = itemNum & Comma & output
                        tempDate = CDate(data.Cells.Item(index - i, DateColumn))
                        i = i + 1
                        previousDate = CDate(data.Cells.Item(index - i, DateColumn))
                    Loop Until previousDate <> tempDate Or remainingCapacity <= 0
                    holiday = data.Cells.Item(index - i, HolidaysColumn)
                    If remainingCapacity <> 0 Or (NoProduction(holiday) And firstRemainingCapacity >= 0) Then
                        Do
                            itemNum = data.Cells.Item(index - i, ItemColumn)
                            i = i + 1
                        Loop Until itemNum <> vbNullString
                    End If
                End If
                ' Format output.
                remainingCapacity = data.Cells.Item(index - i + 1, RemainingCapacityColumn)
                If remainingCapacity < 0 Then output = itemNum & Comma & output
                If Right$(output, Len(Comma)) = Comma Then output = Left$(output, Len(output) - Len(Comma))
                If LenB(output) > 0 Then
                    ProductionFinish = TextToDate(currentDate) & output
                Else
                    ProductionFinish = vbNullString
                End If
            End If
        ElseIf Capacity + lastRemainingCapacity >= 0 And Abs(productAmount - Capacity) < Abs(remainingCapacity) Then
            Do
                itemNum = data.Cells.Item(index - i, ItemColumn)
                i = i + 1
            Loop Until itemNum <> vbNullString
            remainingCapacity = data.Cells.Item(index - i + 1, RemainingCapacityColumn)
            output = itemNum & Comma & output
            If Right$(output, Len(Comma)) = Comma Then output = Left$(output, Len(output) - Len(Comma))
            ProductionFinish = TextToDate(currentDate) + output
        End If
    End If
End Function

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Shows which items were due in repect to their jobs.")
Private Function ShowDueItems(ByVal index As Long, ByVal jobs As Range, ByVal data As Range) As String
Attribute ShowDueItems.VB_Description = "Shows which items were due in repect to their jobs."
    Dim jobData As Object
    Set jobData = CreateObject("Scripting.Dictionary")
    Dim currentRow As Range
    ' Get jobs and their dates.
    For Each currentRow In jobs.Rows
        Dim itemNum As String
        Const JobsItemColumn As Long = 1
        itemNum = currentRow.Cells.Item(1, JobsItemColumn)
        Dim due As Date
        due = currentRow.Cells.Item(1, JobsDueDatesColumn)
        jobData.Item(due) = jobData.Item(due) & itemNum & Comma
    Next
    
    ' Show due jobs at correct dates.
    Dim currentDate As Date
    currentDate = data.Cells.Item(index, DateColumn)
    If jobData.Exists(currentDate) Then
        Dim nextDay As Date
        nextDay = data.Cells.Item(index + 1, DateColumn)
        If nextDay <> currentDate Then
            ShowDueItems = Left$(jobData.Item(currentDate), Len(jobData.Item(currentDate)) - 2)
        Else
            ShowDueItems = vbNullString
        End If
    Else
        ShowDueItems = vbNullString
    End If
End Function
