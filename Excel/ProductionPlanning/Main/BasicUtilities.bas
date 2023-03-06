Attribute VB_Name = "BasicUtilities"
Attribute VB_Description = "Basic utilities for cleaning and protection."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("Basic utilities for cleaning and protection.")
Option Explicit

'@EntryPoint
'@Description("Clears all cells up to a given date.")
Public Sub DeleteUpToDate()
Attribute DeleteUpToDate.VB_Description = "Clears all cells up to a given date."
    ' Variables
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    Dim startingDateCell As Range
    Set startingDateCell = GetCell(ws.Cells, StartingDateRow, StartingDateColumn)
    Dim data As Range
    Dim jobs As Range
    With ws.UsedRange
        Set data = .Range(.Columns.Item(DateColumn), .Columns.Item(SlowdownsColumn))
        Set jobs = .Range(.Columns.Item(JobsDefColumn), .Columns.Item(JobsDueDatesColumn))
    End With
    
    ' Get input.
    Dim inputString As String
    inputString = InputBox(DeletionQuestion, InputLabel, startingDateCell.Value)
    Do Until IsDate(inputString)
        If LenB(inputString) = 0 Then Exit Sub
        inputString = InputBox(NoDateWarning, InputLabel, startingDateCell.Value)
    Loop
    Dim inputDate As Date
    inputDate = CDate(inputString)
    
    ' Delete data up to given date.
    ws.UnProtect
    Dim tempDate As Date
    Dim i As Long
    Dim dateCell As Range
    For i = 0 To Abs(DateDiff("d", inputDate, startingDateCell))
        tempDate = DateAdd("d", -i, inputDate)
        Set dateCell = GetColumn(jobs, JobColumn).Find(tempDate)
        Do Until dateCell Is Nothing
            jobs.Rows.Item(dateCell.row).Delete
            Set dateCell = jobs.Columns.Item(JobColumn).FindNext
        Loop
    Next
    '@Ignore AssignmentNotUsed
    Set dateCell = GetColumn(data, DateColumn).Find(inputDate, GetCell(data, data.Rows.Count - 1, DateColumn), xlValues, SearchDirection:=xlPrevious)
    If Not dateCell Is Nothing Then
        data.Rows.Item(StartingRow & Colon & dateCell.row).Delete
        startingDateCell.Value = DateAdd("d", 1, inputDate)
    End If

    ' Re-protect and confirm success.
    ws.Protect
    MsgBox SlowdownChangeWarning, vbExclamation, WarningLabel
End Sub

' ————————————————————————————————————————————————————— '

'@EntryPoint
'@Description("Toggles protection status of worksheet.")
Public Sub UnProtect()
Attribute UnProtect.VB_Description = "Toggles protection status of worksheet."
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    If ws.ProtectContents Then
        ws.UnProtect
        MsgBox ProtectionLifted
    Else
        ws.Protect
        MsgBox ProtectionEnabled
    End If
End Sub
