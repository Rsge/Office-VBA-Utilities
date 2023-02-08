Attribute VB_Name = "Library"
'@Folder("ProductionPlanning.Main")
Option Explicit

'@Description("Gets the cell in a range at a position.")
Public Function GetCell(ByVal range_ As Range, ByVal row As Long, ByVal column As Long) As Range
Attribute GetCell.VB_Description = "Gets the cell in a range at a position."
    Set GetCell = range_.Cells.Item(row, column)
End Function

'@Description("Gets the column in a range at a position.")
Public Function GetColumn(ByVal range_ As Range, ByVal column As Long) As Range
Attribute GetColumn.VB_Description = "Gets the column in a range at a position."
    Set GetColumn = range_.Columns.Item(column)
End Function

'@Description("Gets the value of a cell in a range at a position.")
Public Function GetCellValue(ByVal range_ As Range, ByVal row As Long, ByVal column As Long) As Variant
Attribute GetCellValue.VB_Description = "Gets the value of a cell in a range at a position."
    GetCellValue = GetCell(range_, row, column).Value
End Function

'@Description("Gets the value of a cell on the active sheet at a position.")
Public Function GetActiveCellValue(ByVal row As Long, ByVal column As Long) As Variant
Attribute GetActiveCellValue.VB_Description = "Gets the value of a cell on the active sheet at a position."
    GetActiveCellValue = GetCell(ActiveSheet.Cells, row, column).Value
End Function

'@Description("Gets a string representing a range of columns from their two column indizes.")
Public Function GetColumnRangeStr(ByVal firstColumn As Long, ByVal secondColumn As Long) As String
Attribute GetColumnRangeStr.VB_Description = "Gets a string representing a range of columns from their two column indizes."
    GetColumnRangeStr = Chr$(firstColumn + ColumnLetterAscii) & Colon & Chr$(secondColumn + ColumnLetterAscii)
End Function
