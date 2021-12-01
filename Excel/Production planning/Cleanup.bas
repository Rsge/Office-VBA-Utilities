Attribute VB_Name = "Cleanup"
Public Sub DeleteUpToDate()
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    Dim StartingDate As Range
    Set StartingDate = WS.Cells(2, 1)
    Dim InputString As String
    Dim Notice As String
    Notice = "Up to which date shall be deleted??" & vbCrLf & "(DD.MM.YYYY)"
    InputString = InputBox(Notice, "Input", StartingDate.Value)
    Do While Not IsDate(InputString)
        If InputString = "" Then Exit Sub
        InputString = InputBox("Input can't be processed as a date." & vbCrLf & vbCrLf & Notice, "Input", StartingDate.Value)
    Loop
    WS.UnProtect
    Dim Data As Range
    Set Data = WS.UsedRange.Columns("A:I")
    Dim InputDate As Date
    InputDate = CDate(InputString)
    Dim DateCell As Range
    
    Dim Jobs As Range
    Set Jobs = WS.UsedRange.Columns("K:L")
    Dim TempDate As Date
    TempDate = InputDate
    Dim i As Integer
    For i = 0 To Abs(DateDiff("d", InputDate, StartingDate))
        TempDate = DateAdd("d", -i, InputDate)
        Set DateCell = Jobs.Columns(2).Find(TempDate)
        Do Until DateCell Is Nothing
            Jobs.Rows(DateCell.Row).Delete
            Set DateCell = Jobs.Columns(2).FindNext
        Loop
        
    Next
    Set DateCell = Data.Columns(1).Find(InputDate, Data.Cells(Data.Rows.Count - 1, 1), xlValues, SearchDirection:=xlPrevious)
    If Not DateCell Is Nothing Then
        Data.Rows("5:" & DateCell.Row).Delete
        StartingDate.Value = DateAdd("d", 1, InputDate)
    End If

    WS.Protect
    Dim Whatever
    Whatever = MsgBox("Please check special slowdown!", vbExclamation, "Warning!")
End Sub

Public Sub UnProtect()
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.ActiveSheet
    Dim Whatever
    If WS.ProtectContents = True Then
        WS.UnProtect
        Whatever = MsgBox("Protection lifted." & vbCrLf & "Changes now possible.")
    Else
        WS.Protect
        Whatever = MsgBox("Protection reestablished.")
    End If
End Sub
