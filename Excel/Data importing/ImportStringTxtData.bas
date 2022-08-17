Attribute VB_Name = "ImportStringTxtData"
Attribute VB_Description = "Importing string data from text file to excel table."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Import"
'@ModuleDescription "Importing string data from text file to excel table."
Option Explicit

'String constants
'@VariableDescription "Path to file with data to import."
Private Const m_path As String = "C:\Test\Example.txt"
Attribute m_path.VB_VarDescription = "Path to file with data to import."

'Int constants
'@VariableDescription "Column
Private Const m_itemColumn As Long = 1
Private Const m_infoColumn As Long = 3


'@EntryPoint
'@Description "Imports string data for items from txt file."
Public Sub ImportStringItemData()
Attribute ImportItemData.VB_Description = "Imports string data for items from txt file."
    'Variables
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    Dim currentLine As String
    Dim itemData() As String
    Dim i As Long
    i = 2
    
    'Get data from file
    Open m_path For Input As fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, currentLine
        itemData = Split(currentLine, vbTab)
        items.Add itemData(0), itemData(1)
    Loop
    
    'Import data to excel table
    Do While LenB(ActiveSheet.Cells(i, m_itemColumn)) <> 0
        itemData(0) = ActiveSheet.Cells(i, m_itemColumn)
        If items.Exists(itemData(0)) Then
            ActiveSheet.Cells(i, m_infoColumn) = items(itemData(0))
            items.Remove (itemData(0))
        End If
        i = i + 1
    Loop
    Dim item As Variant
    For Each item In items.Keys
        ActiveSheet.Cells(i, m_itemColumn) = "'" & item
        ActiveSheet.Cells(i, m_infoColumn) = items(itemData(0))
    Next
End Sub
