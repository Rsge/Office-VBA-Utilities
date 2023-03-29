Attribute VB_Name = "Library"
Attribute VB_Description = "General, often used methods."
'@Folder("DocumentCreation.Base")
'@ModuleDescription("General, often used methods.")
Option Explicit

'@Description("Clears a cell if it only contains spaces. Returns True if cell has changed.")
Public Function RemoveSpaceOnly(ByVal ws As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As Boolean
Attribute RemoveSpaceOnly.VB_Description = "Clears a cell if it only contains spaces. Returns True if cell has changed."
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^ +$"
    If regex.Test(GetCellValue(ws, row_, column_)) Then
        GetCell(ws, row_, column_).Value = vbNullString
        RemoveSpaceOnly = True
    Else
        RemoveSpaceOnly = False
    End If
End Function

'@Description("Checks if a sheet contains an ignore marker. If yes, swaps it with the next sheet and returns True.")
Public Function ChooseCorrectSheet(ByVal wb As Workbook, ByRef ws As Worksheet, ByVal ignore As String) As Boolean
Attribute ChooseCorrectSheet.VB_Description = "Checks if a sheet contains an ignore marker. If yes, swaps it with the next sheet and returns True."
    If Contains(ws.Name, ignore) Then
        wb.Sheets.[_Default](2).Move ws
        Set ws = wb.Sheets.[_Default](1)
        ChooseCorrectSheet = True
    Else
        ChooseCorrectSheet = False
    End If
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

'@Description("Tests if a string is empty.")
Public Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

'@Description("Tests if a string starts with another string.")
Public Function StartsWith(ByVal str As String, ByVal start As String) As Boolean
Attribute StartsWith.VB_Description = "Tests if a string starts with another string."
    StartsWith = InStrB(str, start) = 1
End Function

''@Description("Tests if a string ends with another string.")
'Public Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
'    EndsWith = Right$(str, Len(ending)) = ending
'End Function

'@Description("Tests if a string contains another string.")
Public Function Contains(ByVal str As String, ByVal match As String) As Boolean
Attribute Contains.VB_Description = "Tests if a string contains another string."
    Contains = InStr(str, match) > 0
End Function

'@Description("Tests if a cell has a border on the specified side.")
Public Function HasBorder(ByVal cell As Range, ByVal side As XlBordersIndex) As Boolean
Attribute HasBorder.VB_Description = "Tests if a cell has a border on the specified side."
    HasBorder = cell.Borders.[_Default](side).LineStyle <> xlNone
End Function

''@Description("Removes the first x characters of a string.")
'Public Function RemoveFirst(ByVal str As String, ByVal length As Long) As String
'    RemoveFirst = Right$(str, Len(str) - length)
'End Function

'@Description("Removes the last x characters of a string.")
Public Function RemoveLast(ByVal str As String, ByVal length As Long) As String
Attribute RemoveLast.VB_Description = "Removes the last x characters of a string."
    RemoveLast = Left$(str, Len(str) - length)
End Function

'@Description("Gets the cell on a worksheet at a position.")
Public Function GetCell(ByVal sheet As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As Range
Attribute GetCell.VB_Description = "Gets the cell on a worksheet at a position."
    Set GetCell = sheet.Cells.Item(row_, column_)
End Function

'@Description("Gets the value of a cell on a worksheet at a position.")
Public Function GetCellValue(ByVal sheet As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetCellValue.VB_Description = "Gets the value of a cell on a worksheet at a position."
    GetCellValue = GetCell(sheet, row_, column_).Value
End Function

'@Description("Gets the value of a cell in a row or column at an index.")
Public Function GetCellValueR(ByVal range_ As Range, ByVal index As Long) As Variant
Attribute GetCellValueR.VB_Description = "Gets the value of a cell in a row or column at an index."
    GetCellValueR = range_.Cells.Item(index).Value
End Function

'@Description("Gets the text of a cell in a row or column at an index.")
Public Function GetCellText(ByVal sheet As Worksheet, ByVal row_ As Long, ByVal column_ As Long) As String
Attribute GetCellText.VB_Description = "Gets the text of a cell in a row or column at an index."
    GetCellText = GetCell(sheet, row_, column_).Text
End Function

'@Description("Gets the index of the last used row in a column.")
Public Function GetLastRowIndex(ByVal ws As Worksheet, ByVal column_ As Long) As Long
Attribute GetLastRowIndex.VB_Description = "Gets the index of the last used row in a column."
    GetLastRowIndex = ws.Columns.Item(column_).SpecialCells(xlCellTypeLastCell).Row
End Function

'@Description("Gets the index of the last used column in a row.")
Public Function GetLastColumnIndex(ByVal ws As Worksheet, ByVal row_ As Long) As Long
Attribute GetLastColumnIndex.VB_Description = "Gets the index of the last used column in a row."
    GetLastColumnIndex = ws.Rows.Item(row_).SpecialCells(xlCellTypeLastCell).Column
End Function

''@Description("Gets name of file without the extension.")
'Public Function GetFileNameWithoutExtension(ByVal fileObject As Object) As String
'    GetFileNameWithoutExtension = Split(fileObject.Name, ".")(0)
'End Function

'@Description("Shows an error message box.")
Public Sub ErrBox(ByVal msg As String)
Attribute ErrBox.VB_Description = "Shows an error message box."
    MsgBox msg, vbCritical, ErrorTitle
End Sub

'@Description("Shows a warning message box.")
Public Sub WarnBox(ByVal msg As String)
Attribute WarnBox.VB_Description = "Shows a warning message box."
    MsgBox msg, vbExclamation, WarningTitle
End Sub
