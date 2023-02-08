Attribute VB_Name = "BasicUtilities"
Attribute VB_Description = "Basic utilities for cleaning and protection."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("Basic utilities for cleaning and protection.")
Option Explicit

' String constants
'@VariableDescription("Question about up to which date the calculations should be cleared.")
Private Const m_deletionQuestion As String = "Up to which date shall be deleted?" & vbNewLine & "(DD.MM.YYYY)"
Attribute m_deletionQuestion.VB_VarDescription = "Question about up to which date the calculations should be cleared."
'@VariableDescription("Warning for input not being processable as a date.")
Private Const m_noDateWarning As String = "Input can't be processed as a date." & vbNewLine & vbNewLine & DeletionQuestion
Attribute m_noDateWarning.VB_VarDescription = "Warning for input not being processable as a date."
'@VariableDescription("Warning to check special slowdown after making changes to dates etc.")
Private Const m_slowdownChangeWarning As String = "Please check special slowdown!"
Attribute m_slowdownChangeWarning.VB_VarDescription = "Warning to check special slowdown after making changes to dates etc."
'@VariableDescription("Title of input box to show it needs an input.")
Private Const m_inputLabel As String = "Input"
Attribute m_inputLabel.VB_VarDescription = "Title of input box to show it needs an input."
'@VariableDescription("Title of MsgBox to show it contains a warning.")
Private Const m_warningLabel As String = "Warning!"
Attribute m_warningLabel.VB_VarDescription = "Title of MsgBox to show it contains a warning."
'@VariableDescription("Message for lifted worksheet protection.")
Private Const m_protectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Attribute m_protectionLifted.VB_VarDescription = "Message for lifted worksheet protection."
'@VariableDescription("Message for enforced worksheet protection.")
Private Const m_protectionEnabled As String = "Protection reestablished."
Attribute m_protectionEnabled.VB_VarDescription = "Message for enforced worksheet protection."

' ————————————————————————————————————————————————————— '


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
    Set data = GetColumn(ws.UsedRange, GetColumnRangeStr(DateColumn, SlowdownsColumn))
    Dim jobs As Range
    Set jobs = GetColumn(ws.UsedRange, GetColumnRangeStr(JobsDefColumn, JobsDueDatesColumn))
    
    ' Get input.
    Dim inputString As String
    inputString = InputBox(m_deletionQuestion, m_inputLabel, startingDateCell.Value)
    Do Until IsDate(inputString)
        If LenB(inputString) = 0 Then Exit Sub
        inputString = InputBox(m_noDateWarning, m_inputLabel, startingDateCell.Value)
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
    MsgBox m_slowdownChangeWarning, vbExclamation, m_warningLabel
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
        MsgBox m_protectionLifted
    Else
        ws.Protect
        MsgBox m_protectionEnabled
    End If
End Sub
