Attribute VB_Name = "Library"
Attribute VB_Description = "General, often used methods."
'@Folder("ProductionPlanning.Main")
'@ModuleDescription("General, often used methods.")
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
