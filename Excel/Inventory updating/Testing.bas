Attribute VB_Name = "Testing"
'@Folder("Inventory updating")
Option Explicit

'Excel table column constants (1-based)
''@VariableDescription "Index of items' column in Excel table."
'Private Const m_genItemColumn As Long = 1
''@VariableDescription "Index of amounts' units' column in Excel table."
'Private Const m_genUnitColumn As Long = 2
'@VariableDescription "Index of automatically processed BB-dates' column in Excel table."
Private Const m_autoBBDateColumn As Long = 3
Attribute m_autoBBDateColumn.VB_VarDescription = "Index of automatically processed BB-dates' column in Excel table."
''@VariableDescription "Index of automatically processed old amounts' column in Excel table."
'Private Const m_autoOldAmountColumn As Long = 4
''@VariableDescription "Index of automatically processed amounts' differences' column in Excel table."
'Private Const m_autoDiffAmountColumn As Long = 5
'@VariableDescription "Index of automatically processed new amounts' column in Excel table."
Private Const m_autoNewAmountColumn As Long = 6
Attribute m_autoNewAmountColumn.VB_VarDescription = "Index of automatically processed new amounts' column in Excel table."
'@VariableDescription "Index of automatically processed last change dates' column in Excel table."
Private Const m_autoChangeDateColumn As Long = 7
Attribute m_autoChangeDateColumn.VB_VarDescription = "Index of automatically processed last change dates' column in Excel table."
'@VariableDescription "Index of manually input BB-dates' column in Excel table."
Private Const m_manBBDateColumn As Long = 8
Attribute m_manBBDateColumn.VB_VarDescription = "Index of manually input BB-dates' column in Excel table."
''@VariableDescription "Index of manually input old amounts' column in Excel table."
'Private Const m_manOldAmountColumn As Long = 9
''@VariableDescription "Index of manually input amounts' differences' column in Excel table."
'Private Const m_manDiffAmountColumn As Long = 10
'@VariableDescription "Index of manually input new amounts' column in Excel table."
Private Const m_manNewAmountColumn As Long = 11
Attribute m_manNewAmountColumn.VB_VarDescription = "Index of manually input new amounts' column in Excel table."
'@VariableDescription "Index of manually input last change dates' column in Excel table."
Private Const m_manChangeDateColumn As Long = 12
Attribute m_manChangeDateColumn.VB_VarDescription = "Index of manually input last change dates' column in Excel table."

'@VariableDescription "Date the last import to account for has been done."
Private Const m_checkupDate As String = "12.04.2022"
Attribute m_checkupDate.VB_VarDescription = "Date the last import to account for has been done."
'@VariableDescription "How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations."
Private Const m_diffThresholdPercent As Double = 0.01
Attribute m_diffThresholdPercent.VB_VarDescription = "How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations."

'@EntryPoint
'@Description "Deletes all entries who's automatic and manual last changed dates are before the checkup date."
Public Sub DeleteUnchanged()
Attribute DeleteUnchanged.VB_Description = "Deletes all entries who's automatic and manual last changed dates are before the checkup date."
    Dim i As Long
    i = StartingRow
    Do Until LenB(ActiveSheet.Cells(i, ItemColumn).Value) = 0
        Dim manWasChanged As Boolean
        manWasChanged = ActiveSheet.Cells(i, m_manChangeDateColumn).Value = m_checkupDate
        Dim autoWasChanged As Boolean
        autoWasChanged = ActiveSheet.Cells(i, m_autoChangeDateColumn).Value = m_checkupDate
        If manWasChanged Or autoWasChanged Then
            i = i + 1
        Else
            ActiveSheet.Rows(i).Delete
        End If
    Loop
End Sub

'@EntryPoint
'@Description "Deletes all entries without difference in BB-date and (significant to a threshold) difference in automatic and manual new amount."
Public Sub DeleteEquals()
Attribute DeleteEquals.VB_Description = "Deletes all entries without difference in BB-date and (significant to a threshold) difference in automatic and manual new amount."
    Dim i As Long
    i = StartingRow
    Do Until LenB(ActiveSheet.Cells(i, ItemColumn).Value) = 0
        Dim bbDateMatch As Boolean
        bbDateMatch = CDate(ActiveSheet.Cells(i, m_autoBBDateColumn).Value) = CDate(ActiveSheet.Cells(i, m_manBBDateColumn).Value)
        Dim diff As Double
        diff = Abs(CDbl(ActiveSheet.Cells(i, m_autoNewAmountColumn).Value) - CDbl(ActiveSheet.Cells(i, m_manNewAmountColumn).Value))
        Dim diffThreshold As Double
        diffThreshold = ActiveSheet.Cells(i, m_autoNewAmountColumn).Value * (m_diffThresholdPercent / 100)
        Dim diffMatch As Boolean
        diffMatch = diff < diffThreshold
        If bbDateMatch And diffMatch Then
            ActiveSheet.Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
End Sub
