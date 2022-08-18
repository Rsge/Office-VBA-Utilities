Attribute VB_Name = "ImportTxtData"
Attribute VB_Description = "Importing boolean data from text file to excel table."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Import"
'@ModuleDescription "Importing boolean data from text file to excel table."
Option Explicit

'String constants
'@VariableDescription "Path to file with data to import."
Private Const m_path As String = "C:\Test\Example.txt"
Attribute m_path.VB_VarDescription = "Path to file with data to import."
'@VariableDescription "Label for item missing an entry in the data file."
Private Const m_hasNothing As String = "None"
Attribute m_hasNothing.VB_VarDescription = "Label for item missing an entry in the data file."
'@VariableDescription "Label for item missing specific info."
Private Const m_hasNoSpecific As String = "Specific"
Attribute m_hasNoSpecific.VB_VarDescription = "Label for item missing specific info."
'@VariableDescription "Label for item missing an entry in the Excel sheet."
Private Const m_hasNoEntry As String = "Missing Entry"
Attribute m_hasNoEntry.VB_VarDescription = "Label for item missing an entry in the Excel sheet."

'Int constants
'@VariableDescription "Column
Private Const m_itemColumn As Long = 1
Private Const m_infoColumn As Long = 5


'@EntryPoint
'@Description "Imports boolean data for items from txt file."
Public Sub ImportBooleanItemData()
Attribute ImportItemData.VB_Description = "Imports boolean data for items from txt file."
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
    
    'Import data to Excel table
    Do While LenB(ActiveSheet.Cells(i, m_itemColumn)) <> 0
        itemData(0) = ActiveSheet.Cells(i, m_itemColumn)
        If items.Exists(itemData(0)) Then
            If Not CBool(items(itemData(0))) Then
                If InStrB(ActiveSheet.Cells(i, m_infoColumn), m_hasNoEntry) > 0 Then
                    ActiveSheet.Cells(i, m_infoColumn) = m_hasNoEntry & " & " & m_hasNoSpecific
                Else
                    ActiveSheet.Cells(i, m_infoColumn) = m_hasNoSpecific
                End If
            End If
            items.Remove (itemData(0))
        Else
            ActiveSheet.Cells(i, m_infoColumn) = m_hasNothing
        End If
        i = i + 1
    Loop
    Dim item As Variant
    For Each item In items.Keys
        ActiveSheet.Cells(i, m_itemColumn) = "'" & item
        ActiveSheet.Cells(i, m_infoColumn) = m_hasNoEntry
    Next
End Sub
