Attribute VB_Name = "ItemStatusChecker"
Attribute VB_Description = "Reading string data from text file to check entries in excel table."
'@Folder "Import"
'@ModuleDescription "Reading string data from text file to check entries in excel table."
Option Explicit

'String constants
'@VariableDescription "Path to file with data to import."
Private Const m_path As String = "C:\Test\Example.txt"
Attribute m_path.VB_VarDescription = "Path to file with data to import."
'@VariableDescription "String to put in info column if item exists in file."
Private Const m_exists As String = "Yes"
Attribute m_exists.VB_VarDescription = "String to put in info column if item exists in file."

'Int constants
'@VariableDescription "The column containing the item number."
Private Const m_itemColumn As Long = 1
Attribute m_itemColumn.VB_VarDescription = "The column containing the item number."
'@VariableDescription "The column to input the validation result."
Private Const m_infoColumn As Long = 2
Attribute m_infoColumn.VB_VarDescription = "The column to input the validation result."
'@VariableDescription "The frist row in excel table with data."
Private Const m_startingRow As Long = 4
Attribute m_startingRow.VB_VarDescription = "The frist row in excel table with data."


'@EntryPoint
'@Description "Imports string data for items from txt file."
Public Sub CheckItemStatus()
Attribute CheckItemStatus.VB_Description = "Imports string data for items from txt file."
    'Variables
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim items As Object
    Set items = CreateObject("System.Collections.ArrayList")
    Dim currentLine As String
    Dim currentItem As String
    Dim i As Long
    i = m_startingRow

    'Get data from file
    Open m_path For Input As fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, currentLine
        items.Add currentLine
    Loop
    Close fileNumber

    'Import data to excel table
    Do While LenB(ActiveSheet.Cells(i, m_itemColumn)) <> 0
        currentItem = ActiveSheet.Cells(i, m_itemColumn)
        If items.Contains(currentItem) Then
            ActiveSheet.Cells(i, m_infoColumn) = m_exists
            items.Remove (currentItem)
        End If
        i = i + 1
    Loop
    'For updating import after change comment out this part
    Dim item As Variant
    For Each item In items
        ActiveSheet.Cells(i, m_itemColumn) = "'" & item
        ActiveSheet.Cells(i, m_infoColumn) = m_exists
    Next
End Sub
