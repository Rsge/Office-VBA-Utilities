Attribute VB_Name = "Library"
Attribute VB_Description = "General, often used methods."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("General, often used methods.")
Option Explicit

'@Description("Tests if a string is empty.")
Public Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

'@Description("Gets the cell in a range at a position.")
Public Function GetCell(ByVal range_ As Range, ByVal row_ As Long, ByVal column_ As Long) As Range
Attribute GetCell.VB_Description = "Gets the cell in a range at a position."
    Set GetCell = range_.Cells.Item(row_, column_)
End Function

'@Description("Gets the column in a range at a position.")
Public Function GetColumn(ByVal range_ As Range, ByVal column_ As Long) As Range
Attribute GetColumn.VB_Description = "Gets the column in a range at a position."
    Set GetColumn = range_.Columns.Item(column_)
End Function

'@Description("Gets the value of a cell in a range at a position.")
Public Function GetCellValue(ByVal range_ As Range, ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetCellValue.VB_Description = "Gets the value of a cell in a range at a position."
    GetCellValue = GetCell(range_, row_, column_).Value
End Function

'@Description("Gets the value of a cell on the active sheet at a position.")
Public Function GetActCellValue(ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetActCellValue.VB_Description = "Gets the value of a cell on the active sheet at a position."
    GetActCellValue = GetCell(ActiveSheet.Cells, row_, column_).Value
End Function
