Attribute VB_Name = "Legacy"
Attribute VB_Description = "Module for saving legacy code."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Module for saving legacy code."
Option Explicit

'String constant
Private Const DatePrefix As String = "On "


'Date to string functions
'@Description "Converts a date to a string in default european date format.
Private Function TextCurrentDate(ByVal CurrentDate As Date) As String
    TextCurrentDate = Format$(CurrentDate, "dd.mm.yyyy")
End Function
'@Description "Returns an info text with the given date."
Private Function TextToDate(ByVal CurrentDate As Date) As String
Attribute TextToDate.VB_Description = "Returns an info text with the given date."
    TextToDate = DatePrefix & TextCurrentDate(CurrentDate) & Colon & Space(1)
End Function


'@EntryPoint
'@Description "Calculated the point at which the production of an item would finish"
Private Function ProductionFinish(ByVal Capacity As Long, ByVal Index As Long, ByVal Data As Range) As String
Attribute ProductionFinish.VB_Description = "Calculated the point at which the production of an item would finish"
    'Variables
    Dim CurrentDate As Date
    CurrentDate = CDate(Data.Cells.Item(Index, DateColumn))
    Dim ProductAmount As Long
    ProductAmount = Data.Cells.Item(Index, AmountColumn)
    Dim RemainingCapacity As Long
    RemainingCapacity = Data.Cells.Item(Index, RemainingCapacityColumn)
    Dim Holiday As String
    Holiday = Data.Cells.Item(Index, HolidaysColumn)
    Dim Item As String
    Item = Data.Cells.Item(Index, ItemColumn)
    Dim i As Long
    i = 1
    Dim Output As String
    Dim TempDate As Date
    
    'If no production is done or there is remaining capacity, item isn't done,
    'else if it's the first entry, item is only done when there's no remaining capacity,
    'else it has to be further calculated
    If LenB(Holiday) = 0 Or RemainingCapacity > 0 Then
        ProductionFinish = vbNullString
    ElseIf Index = 1 Then
        If RemainingCapacity = 0 Then
            ProductionFinish = TextToDate(CurrentDate) & Item
        Else
            ProductionFinish = vbNullString
        End If
    Else
        Dim LastRemainingCapacity As Long
        LastRemainingCapacity = Data.Cells.Item(Index - 1, RemainingCapacityColumn)
        Dim FirstRemainingCapacity As Long
        FirstRemainingCapacity = RemainingCapacity
        Dim PreviousDate As Date
        PreviousDate = CDate(Data.Cells.Item(Index - 1, DateColumn))
        'If there's enough capacity to finish the current item, it can be added as finished
        If Capacity + LastRemainingCapacity >= 0 And ProductAmount <= Capacity Then
            If ProductAmount = 0 And RemainingCapacity < -Capacity Then
                ProductionFinish = vbNullString
            Else
                'If there's remaining capacity, the current item is finished
                If RemainingCapacity >= 0 Then Output = Item
                'And if there's enough capacity, older items will be finished now, too
                If ProductAmount <> Capacity Then
                    Do
                        Item = Data.Cells.Item(Index - i, ItemColumn)
                        RemainingCapacity = Data.Item(Index - i, RemainingCapacityColumn)
                        If Item <> vbNullString And RemainingCapacity <> 0 Then Output = Item & Comma & Output
                        TempDate = CDate(Data.Cells.Item(Index - i, DateColumn))
                        i = i + 1
                        PreviousDate = CDate(Data.Cells.Item(Index - i, DateColumn))
                    Loop Until PreviousDate <> TempDate Or RemainingCapacity <= 0
                    Holiday = Data.Cells.Item(Index - i, HolidaysColumn)
                    If RemainingCapacity <> 0 Or (NoProduction(Holiday) And FirstRemainingCapacity >= 0) Then
                        Do
                            Item = Data.Cells.Item(Index - i, ItemColumn)
                            i = i + 1
                        Loop Until Item <> vbNullString
                    End If
                End If
                'Format output
                RemainingCapacity = Data.Cells.Item(Index - i + 1, RemainingCapacityColumn)
                If RemainingCapacity < 0 Then Output = Item & Comma & Output
                If Right$(Output, Len(Comma)) = Comma Then Output = Left$(Output, Len(Output) - Len(Comma))
                If LenB(Output) > 0 Then
                    ProductionFinish = TextToDate(CurrentDate) & Output
                Else
                    ProductionFinish = vbNullString
                End If
            End If
        ElseIf Capacity + LastRemainingCapacity >= 0 And Abs(ProductAmount - Capacity) < Abs(RemainingCapacity) Then
            Do
                Item = Data.Cells.Item(Index - i, ItemColumn)
                i = i + 1
            Loop Until Item <> vbNullString
            RemainingCapacity = Data.Cells.Item(Index - i + 1, RemainingCapacityColumn)
            Output = Item & Comma & Output
            If Right$(Output, Len(Comma)) = Comma Then Output = Left$(Output, Len(Output) - Len(Comma))
            ProductionFinish = TextToDate(CurrentDate) + Output
        End If
    End If
End Function

'@EntryPoint
'@Description "Showed which items were due in repect to their jobs."
Private Function ShowDueItems(ByVal Index As Long, ByVal Jobs As Range, ByVal Data As Range) As String
Attribute ShowDueItems.VB_Description = "Showed which items were due in repect to their jobs."
    Dim JobData As Object
    Set JobData = CreateObject("Scripting.Dictionary")
    Dim CurrentRow As Range
    'Get jobs and their dates
    For Each CurrentRow In Jobs.Rows
        Dim Item As String
        Const JobsItemColumn As Long = 1
        Item = CurrentRow.Cells.Item(1, JobsItemColumn)
        Dim Due As Date
        Due = CurrentRow.Cells.Item(1, JobsDueDatesColumn)
        JobData(Due) = JobData(Due) & Item & Comma
    Next
    
    'Show due jobs at correct dates
    Dim CurrentDate As Date
    CurrentDate = Data.Cells.Item(Index, DateColumn)
    If JobData.Exists(CurrentDate) Then
        Dim NextDay As Date
        NextDay = Data.Cells.Item(Index + 1, DateColumn)
        If NextDay <> CurrentDate Then
            ShowDueItems = Left$(JobData(CurrentDate), Len(JobData(CurrentDate)) - 2)
        Else
            ShowDueItems = vbNullString
        End If
    Else
        ShowDueItems = vbNullString
    End If
End Function
