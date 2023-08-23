Attribute VB_Name = "BooleanDataImport"
Attribute VB_Description = "Imports boolean data from text file to Excel table."
'@Folder("Imports")
'@ModuleDescription("Imports boolean data from text file to Excel table.")
Option Explicit

' String constants
'@VariableDescription("Path to file with data to import.")
Private Const m_path As String = "C:\Test\Example.txt"
Attribute m_path.VB_VarDescription = "Path to file with data to import."
'@VariableDescription("Label for item missing an entry in the data file.")
Private Const m_hasNothing As String = "None"
Attribute m_hasNothing.VB_VarDescription = "Label for item missing an entry in the data file."
'@VariableDescription("Label for item missing packing info.")
Private Const m_hasNoSpecific As String = "Specific"
Attribute m_hasNoSpecific.VB_VarDescription = "Label for item missing packing info."
'@VariableDescription("Label for item missing an entry in the Excel sheet.")
Private Const m_hasNoEntry As String = "Missing Entry"
Attribute m_hasNoEntry.VB_VarDescription = "Label for item missing an entry in the Excel sheet."

' Integer constants
'@VariableDescription("The frist row in Excel table with data.")
Private Const m_startingRow As Long = 2
Attribute m_startingRow.VB_VarDescription = "The frist row in Excel table with data."
'@VariableDescription("The column containing the item number.")
Private Const m_itemColumn As Long = 1
Attribute m_itemColumn.VB_VarDescription = "The column containing the item number."
'@VariableDescription("The column to import to from file's boolean data.")
Private Const m_infoColumn As Long = 5
Attribute m_infoColumn.VB_VarDescription = "The column to import to from file's boolean data."

' Boolean constants
'@VariableDescription("Wether to just update the item list or add new items at the end, too.")
Private Const m_updateOnly As Boolean = True
Attribute m_updateOnly.VB_VarDescription = "Wether to just update the item list or add new items at the end, too."

' ————————————————————————————————————————————————————— '

'@Description("Gets the cell on the active worksheet at a position.")
Public Function GetActCell(ByVal row_ As Long, ByVal column_ As Long) As Range
Attribute GetActCell.VB_Description = "Gets the cell on the active worksheet at a position."
    Set GetActCell = ActiveSheet.Cells.Item(row_, column_)
End Function

'@Description("Gets the value of a cell on the active worksheet at a position.")
Public Function GetActCellValue(ByVal row_ As Long, ByVal column_ As Long) As Variant
Attribute GetActCellValue.VB_Description = "Gets the value of a cell on the active worksheet at a position."
    GetActCellValue = GetActCell(row_, column_).Value
End Function

'@Description("Sets the value of a cell on the active worksheet at a position to a value.")
Public Sub SetActCellValue(ByVal row_ As Long, ByVal column_ As Long, ByVal val As Variant)
Attribute SetActCellValue.VB_Description = "Sets the value of a cell on the active worksheet at a position to a value."
    GetActCell(row_, column_).Value = val
End Sub

'@EntryPoint
'@Description("Imports boolean data for items from txt file.")
Public Sub ImportBooleanItemData()
Attribute ImportBooleanItemData.VB_Description = "Imports boolean data for items from txt file."
    ' Get data from file.
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    Dim currentLine As String
    Dim itemData() As String
    Open m_path For Input As fileNumber
    Do Until EOF(fileNumber)
        Line Input #fileNumber, currentLine
        itemData = Split(currentLine, vbTab)
        items.Add itemData(0), itemData(1)
    Loop
    Close fileNumber
    
    ' Import data to Excel table.
    Dim i As Long
    i = m_startingRow
    Dim infoCell As Range
    Do Until LenB(GetActCellValue(i, m_itemColumn)) = 0
        itemData(0) = GetActCellValue(i, m_itemColumn)
        Set infoCell = GetActCell(i, m_infoColumn)
        If items.Exists(itemData(0)) Then
            If Not CBool(items.Item(itemData(0))) Then
                If InStr(infoCell.Value, m_hasNoEntry) > 0 Then
                    infoCell.Value = m_hasNoEntry & " & " & m_hasNoPacking
                Else
                    infoCell.Value = m_hasNoPacking
                End If
            End If
            items.Remove (itemData(0))
        Else
            infoCell.Value = m_hasNothing
        End If
        i = i + 1
    Loop
    ' Add new items at end of table.
    If Not m_updateOnly Then
        Dim item_ As Variant
        For Each item_ In items.Keys
            SetActCellValue i, m_itemColumn, "'" & item_
            SetActCellValue i, m_infoColumn, m_hasNoEntry
            i = i + 1
        Next
    End If
End Sub
