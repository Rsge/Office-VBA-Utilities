Attribute VB_Name = "Cleanup"
Attribute VB_Description = "Module for cleanup utils."
'@Folder "Production planning"
'@ModuleDescription "Module for cleanup utils."
Option Explicit

'String constants
'@VariableDescription "Question about up to which date the calculations should be cleared."
Private Const DeletionQuestion As String = "Up to which date shall be deleted?" & vbNewLine & "(DD.MM.YYYY)"
Attribute DeletionQuestion.VB_VarDescription = "Question about up to which date the calculations should be cleared."
'@VariableDescription "Warning for input not being processable as a date."
Private Const NoDateWarning As String = "Input can't be processed as a date." & vbNewLine & vbNewLine & DeletionQuestion
Attribute NoDateWarning.VB_VarDescription = "Warning for input not being processable as a date."
'@VariableDescription "Warning to check special slowdown after making changes to dates etc."
Private Const SlowdownChangeWarning As String = "Please check special slowdown!"
Attribute SlowdownChangeWarning.VB_VarDescription = "Warning to check special slowdown after making changes to dates etc."
'@VariableDescription "Title of input box to show it needs an input."
Private Const InputLabel As String = "Input"
Attribute InputLabel.VB_VarDescription = "Title of input box to show it needs an input."
'@VariableDescription "Title of MsgBox to show it contains a warning."
Private Const WarningLabel As String = "Warning!"
Attribute WarningLabel.VB_VarDescription = "Title of MsgBox to show it contains a warning."
'@VariableDescription "Message for lifted worksheet protection."
Private Const ProtectionLifted As String = "Protection lifted." & vbNewLine & "Changes now possible."
Attribute ProtectionLifted.VB_VarDescription = "Message for lifted worksheet protection."
'@VariableDescription "Message for enforced worksheet protection."
Private Const ProtectionEnabled As String = "Protection reestablished."
Attribute ProtectionEnabled.VB_VarDescription = "Message for enforced worksheet protection."


'@EntryPoint
'@Description "Clear all cells up to a given date."
Public Sub DeleteUpToDate()
Attribute DeleteUpToDate.VB_Description = "Clear all cells up to a given date."
    'Variables
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    Dim StartingDateCell As Range
    Set StartingDateCell = WS.Cells.Item(StartingDateRow, StartingDateColumn)
    Dim Data As Range
    Set Data = WS.UsedRange.Columns.Item(Chr$(DateColumn + ColumnLetterAscii) & Colon & Chr$(SlowdownsColumn + ColumnLetterAscii))
    Dim Jobs As Range
    Set Jobs = WS.UsedRange.Columns.Item(Chr$(JobsDefColumn + ColumnLetterAscii) & Colon & Chr$(JobsDueDatesColumn + ColumnLetterAscii))
    
    'Get input
    Dim InputString As String
    InputString = InputBox(DeletionQuestion, InputLabel, StartingDateCell.Value)
    Do While Not IsDate(InputString)
        If LenB(InputString) = 0 Then Exit Sub
        InputString = InputBox(NoDateWarning, InputLabel, StartingDateCell.Value)
    Loop
    Dim InputDate As Date
    InputDate = CDate(InputString)
    
    'Delete data up to given date
    WS.UnProtect
    Dim TempDate As Date
    Dim i As Long
    Dim DateCell As Range
    For i = 0 To Abs(DateDiff("d", InputDate, StartingDateCell))
        TempDate = DateAdd("d", -i, InputDate)
        Set DateCell = Jobs.Columns.Item(2).Find(TempDate)
        Do Until DateCell Is Nothing
            Jobs.Rows.Item(DateCell.Row).Delete
            Set DateCell = Jobs.Columns.Item(2).FindNext
        Loop
        
    Next
    '@Ignore AssignmentNotUsed
    Set DateCell = Data.Columns.Item(1).Find(InputDate, Data.Cells.Item(Data.Rows.Count - 1, DateColumn), xlValues, SearchDirection:=xlPrevious)
    If Not DateCell Is Nothing Then
        Data.Rows.Item(StartingRow & Colon & DateCell.Row).Delete
        StartingDateCell.Value = DateAdd("d", 1, InputDate)
    End If

    'Re-protect and confirm success
    WS.Protect
    '@Ignore VariableNotUsed
    Dim Whatever As VbMsgBoxResult
    '@Ignore AssignmentNotUsed
    Whatever = MsgBox(SlowdownChangeWarning, vbExclamation, WarningLabel)
End Sub

'@EntryPoint
'@Description "Toggle protection status of worksheet."
Public Sub UnProtect()
Attribute UnProtect.VB_Description = "Toggle protection status of worksheet."
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    '@Ignore VariableNotUsed
    Dim Whatever As VbMsgBoxResult
    If WS.ProtectContents = True Then
        WS.UnProtect
        Whatever = MsgBox(ProtectionLifted)
    Else
        WS.Protect
        Whatever = MsgBox(ProtectionEnabled)
    End If
End Sub
