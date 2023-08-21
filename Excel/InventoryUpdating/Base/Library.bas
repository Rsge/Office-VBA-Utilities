Attribute VB_Name = "Library"
Attribute VB_Description = "General, useful methods."
'@Folder("InventoryUpdating.Base")
'@ModuleDescription("General, useful methods.")
Option Explicit

'@Description("Tests if a string starts with another string.")
Public Function StartsWith(ByVal str As String, ByVal start As String) As Boolean
Attribute StartsWith.VB_Description = "Tests if a string starts with another string."
    StartsWith = InStrB(str, start) = 1
End Function

'@Description("Tests if a string ends with another string.")
Public Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
Attribute EndsWith.VB_Description = "Tests if a string ends with another string."
    EndsWith = Right$(str, Len(ending)) = ending
End Function

'@Description("Tests if a string contains another string.")
Public Function Contains(ByVal str As String, ByVal match As String) As Boolean
Attribute Contains.VB_Description = "Tests if a string contains another string."
    Contains = InStr(str, match) > 0
End Function

'@Description("Tests if a string is empty.")
Public Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

' ————————————————————————————————————————————————————— '

'@Description("Gets the cell on a worksheet at a position.")
Public Function GetCell(ByVal ws As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As Range
Attribute GetCell.VB_Description = "Gets the cell on a worksheet at a position."
    Set GetCell = ws.Cells.Item(row_, column_)
End Function

'@Description("Gets the cell on the active worksheet at a position.")
Public Function GetActCell(ByVal row_ As Long, ByVal column_ As Long) As Range
Attribute GetActCell.VB_Description = "Gets the cell on the active worksheet at a position."
    Set GetActCell = GetCell(ActiveSheet, row_, column_)
End Function

'@Description("Gets the value of a cell on the active worksheet at a position.")
Public Function GetActCellValue(ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetActCellValue.VB_Description = "Gets the value of a cell on the active worksheet at a position."
    GetActCellValue = GetActCell(row_, column_).Value
End Function

'@Description("Sets the value of a cell on the active worksheet at a position to a value.")
Public Sub SetActCellValue(ByVal row_ As Long, ByVal column_ As Long, ByVal val As Variant)
Attribute SetActCellValue.VB_Description = "Sets the value of a cell on the active worksheet at a position to a string."
    GetActCell(row_, column_).Value = val
End Sub


'@Description("Inserts a new row at given position on sheet.")
Public Sub CreateNewRow(ByVal ws As Worksheet, ByVal row_ As Long, ByVal copyFrom As Long)
Attribute CreateNewRow.VB_Description = "Inserts a new row at given position on sheet."
    With ws.Rows
        .Item(row_).Insert
        .Item(row_ + copyFrom).Copy .Item(row_)
    End With
End Sub

'@Description("Inserts a new row at given position on active sheer.")
Public Sub CreateNewActRow(ByVal row_ As Long, ByVal copyFrom As Long)
Attribute CreateNewActRow.VB_Description = "Inserts a new row at given position on active sheer."
    CreateNewRow ActiveSheet, row_, copyFrom
End Sub

' ————————————————————————————————————————————————————— '

'@Description("Shows a MsgBox in a warning style.")
Public Sub WarnBox(ByVal msg As String)
Attribute WarnBox.VB_Description = "Shows a MsgBox in a warning style."
    MsgBox msg, vbExclamation, WarnBoxTitle
End Sub

'@Description("Shows an Ok-Cancel-MsgBox and returns True is Cancel was pressed.")
Public Function MsgBoxCanceled(ByVal msg As String) As Boolean
Attribute MsgBoxCanceled.VB_Description = "Shows an Ok-Cancel-MsgBox and returns True is Cancel was pressed."
    MsgBoxCanceled = MsgBox(msg, vbOKCancel) = vbCancel
End Function

' ————————————————————————————————————————————————————— '

'@Description("Returns if the active Workbook contains a sheet with a name starting with a specified string.")
Public Function ContainsSheetStartingWith(ByVal str As String) As Boolean
Attribute ContainsSheetStartingWith.VB_Description = "Returns if the active Workbook contains a sheet with a name starting with a specified string."
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If StartsWith(ws.Name, str) Then
            ContainsSheetStartingWith = True
            Exit Function
        End If
    Next
    ContainsSheetStartingWith = False
End Function

'@Description("Deletes all sheets in active Workbook which's name starts with a specified string.")
Public Sub DeleteSheetsStartingWith(ByVal str As String)
Attribute DeleteSheetsStartingWith.VB_Description = "Deletes all sheets in active Workbook which's name starts with a specified string."
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Sheets
        If StartsWith(ws.Name, str) Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' ————————————————————————————————————————————————————— '

'@Description("Gets name of file without the extension.")
Public Function GetFileNameWithoutExtension(ByVal fileObject As Object) As String
Attribute GetFileNameWithoutExtension.VB_Description = "Gets name of file without the extension."
    GetFileNameWithoutExtension = Split(fileObject.Name, ".")(0)
End Function

'@Description("Replaces occurences of {} in a string with the specified replacements.")
Public Function FormatString(ByVal str As String, ParamArray replacements() As Variant) As String
Attribute FormatString.VB_Description = "Replaces occurences of {} in a string with the specified replacements."
    Dim parts() As String
    parts = Split(str, "{}")
    Dim out As Object
    Set out = CreateObject("System.Collections.ArrayList")
    Dim i As Long
    i = 0
    Dim replacement As Variant
    For Each replacement In replacements
        out.Add parts(i)
        out.Add replacement
        i = i + 1
    Next
    out.Add parts(i)
    FormatString = Join(out.ToArray, vbNullString)
End Function

'@Description("Gets number of last lines of file at path (default 1).")
Public Function GetFirstLine(ByVal filePath As String, Optional ByVal lineCount As Long = 1) As String()
Attribute GetFirstLine.VB_Description = "Gets number of last lines of file at path (default 1)."
    Dim firstLines() As String
    ReDim firstLines(0 To lineCount - 1)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileStream As Object
    Set fileStream = fso.OpenTextFile(filePath, 1, False)
    Dim i As Long

    ' Open file.
    For i = 0 To lineCount - 1
        If fileStream.AtEndOfStream Then Exit For
        firstLines(i) = fileStream.ReadLine
    Next
    fileStream.Close
    GetFirstLine = firstLines
End Function

'@Description("Gets number of last lines of file at path (default 1).")
Public Function GetLastLine(ByVal filePath As String, Optional ByVal lineCount As Long = 1) As String()
Attribute GetLastLine.VB_Description = "Gets number of last lines of file at path (default 1)."
    Dim fileNumber As Long
    ' Use first unused file number.
    fileNumber = FreeFile
    Dim pointer As Long
    ' Create string of fixed length 1.
    Dim char As String * 1
    Dim currentLineNumber As Long
    currentLineNumber = 0
    Dim lastLines() As String
    ReDim lastLines(0 To lineCount - 1)

    ' Open file.
    Open filePath For Binary As fileNumber
    ' Set pointer to last position in file.
    pointer = LOF(fileNumber)
    Do
        ' Read char at position "pointer" into "char".
        Get fileNumber, pointer, char
        If char = vbCr Then
            ' Simply skip CRs for Linux compat.
            pointer = pointer - 1
        ElseIf char = vbLf Then
            ' Read Count last lines of file.
            If currentLineNumber < lineCount - 1 Then
                currentLineNumber = currentLineNumber + 1
                pointer = pointer - 1
            Else
                Exit Do
            End If
        Else
            pointer = pointer - 1
            ' Add char to result string.
            lastLines(currentLineNumber) = char & lastLines(currentLineNumber)
        End If
    Loop
    Close fileNumber
    GetLastLine = lastLines
End Function

'@Description("Builds file name of a file based on active file's name and current date with specified date format.")
Public Function BuildWBName(ByVal dateFormat As String, Optional ByVal isMakroWB As Boolean = False) As String
Attribute BuildWBName.VB_Description = "Builds file name of a file based on active file's name and current date with specified date format."
    Dim fileRegex As Object
    Set fileRegex = CreateObject("VBScript.RegExp")
    fileRegex.Pattern = FileNameFormatRegex
    Dim matches As Object
    Set matches = fileRegex.Execute(ThisWorkbook.Name)
    Dim match As Object
    Dim baseName As String
    For Each match In matches
        ' There should always only be one match.
        baseName = match.Submatches.Item(0)
    Next
    Dim ext As String
    If isMakroWB Then
        ext = ".xlsm"
    Else
        ext = ".xlsx"
    End If
    BuildWBName = baseName & Format$(Now, dateFormat) & ext
End Function

' ————————————————————————————————————————————————————— '

'@Description("Creates a list from a one-column-table as range.")
Public Function GetTableAsList(ByVal table As Range) As Object
Attribute GetTableAsList.VB_Description = "Creates a list from a one-column-table as range."
    If table.Columns.Count <> 1 Then Err.Raise 1
    Dim list As Object
    Set list = CreateObject("System.Collections.ArrayList")
    Dim cell As Range
    For Each cell In table.Cells
        list.Add CStr(cell.Value)
    Next
    Set GetTableAsList = list
End Function
