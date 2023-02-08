Attribute VB_Name = "StringDataImport"
Attribute VB_Description = "Imports string data from text file to Excel table."
'@Folder("Imports")
'@ModuleDescription("Imports string data from text file to Excel table.")
Option Explicit

' String constants
'@VariableDescription("Path to file with data to import.")
Private Const m_path As String = "C:\Test\Example.txt"
Attribute m_path.VB_VarDescription = "Path to file with data to import."

' Integer constants
'@VariableDescription("The column containing the item number.")
Private Const m_itemColumn As Long = 1
Attribute m_itemColumn.VB_VarDescription = "The column containing the item number."
'@VariableDescription("The column to import to from file's string data.")
Private Const m_infoColumn As Long = 3
Attribute m_infoColumn.VB_VarDescription = "The column to import to from file's string data."
'@VariableDescription("The frist row in Excel table with data.")
Private Const m_startingRow As Long = 2
Attribute m_startingRow.VB_VarDescription = "The frist row in Excel table with data."

' Boolean constants
'@VariableDescription("Wether to just update the item list or add new items at the end, too.")
Private Const m_updateOnly As Boolean = True
Attribute m_updateOnly.VB_VarDescription = "Wether to just update the item list or add new items at the end, too."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Imports string data for items from txt file.")
Public Sub ImportStringItemData()
Attribute ImportStringItemData.VB_Description = "Imports string data for items from txt file."
    ' Variables
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    Dim currentLine As String
    Dim itemData() As String
    Dim i As Long
    i = m_startingRow
    Dim infoCell As Range
    
    ' Get data from file.
    Open m_path For Input As fileNumber
    Do Until EOF(fileNumber)
        Line Input #fileNumber, currentLine
        itemData = Split(currentLine, vbTab)
        items.Add itemData(0), itemData(1)
    Loop
    Close fileNumber
    
    ' Import data to Excel table.
    Do Until LenB(ActiveSheet.Cells(i, m_itemColumn).Value) = 0
        itemData(0) = ActiveSheet.Cells(i, m_itemColumn).Value
        Set infoCell = ActiveSheet.Cells.Item(i, m_infoColumn)
        If items.Exists(itemData(0)) And LenB(infoCell.Value) = 0 Then
            infoCell.Value = items.Item(itemData(0))
            items.Remove (itemData(0))
        End If
        i = i + 1
    Loop
    ' Add new items at end of table.
    If Not m_updateOnly Then
        Dim item_ As Variant
        For Each item_ In items.Keys
            ActiveSheet.Cells.Item(i, m_itemColumn).Value = "'" & item_
            ActiveSheet.Cells.Item(i, m_infoColumn).Value = items.Item(itemData(0))
            i = i + 1
        Next
    End If
End Sub
