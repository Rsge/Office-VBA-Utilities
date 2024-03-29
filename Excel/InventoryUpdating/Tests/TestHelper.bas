Attribute VB_Name = "TestHelper"
Attribute VB_Description = "Helps in testing if automatic import works as intended."
'@Folder("InventoryUpdating.Tests")
'@ModuleDescription("Helps in testing if automatic import works as intended.")
Option Explicit

' Runtime constant
'@VariableDescription("Date the last import to account for in testing has been done.")
Private Const m_checkupDate As String = "01.01.2000"
Attribute m_checkupDate.VB_VarDescription = "Date the last import to account for in testing has been done."

' Excel table column constants (1-based)
''@VariableDescription("Index of items' column in Excel table.")
'Private Const m_genItemColumn As Long = 1
''@VariableDescription("Index of amounts' units' column in Excel table.")
'Private Const m_genUnitColumn As Long = 2
'@VariableDescription("Index of automatically processed BB-dates' column in Excel table.")
Private Const m_autoBBDateColumn As Long = 3
Attribute m_autoBBDateColumn.VB_VarDescription = "Index of automatically processed BB-dates' column in Excel table."
''@VariableDescription("Index of automatically processed old amounts' column in Excel table.")
'Private Const m_autoOldAmountColumn As Long = 4
''@VariableDescription("Index of automatically processed amounts' differences' column in Excel table.")
'Private Const m_autoDiffAmountColumn As Long = 5
'@VariableDescription("Index of automatically processed new amounts' column in Excel table.")
Private Const m_autoNewAmountColumn As Long = 6
Attribute m_autoNewAmountColumn.VB_VarDescription = "Index of automatically processed new amounts' column in Excel table."
'@VariableDescription("Index of automatically processed last change dates' column in Excel table.")
Private Const m_autoChangeDateColumn As Long = 7
Attribute m_autoChangeDateColumn.VB_VarDescription = "Index of automatically processed last change dates' column in Excel table."
'@VariableDescription("Index of manually input BB-dates' column in Excel table.")
Private Const m_manBBDateColumn As Long = 8
Attribute m_manBBDateColumn.VB_VarDescription = "Index of manually input BB-dates' column in Excel table."
''@VariableDescription("Index of manually input old amounts' column in Excel table.")
'Private Const m_manOldAmountColumn As Long = 9
''@VariableDescription("Index of manually input amounts' differences' column in Excel table.")
'Private Const m_manDiffAmountColumn As Long = 10
'@VariableDescription("Index of manually input new amounts' column in Excel table.")
Private Const m_manNewAmountColumn As Long = 11
Attribute m_manNewAmountColumn.VB_VarDescription = "Index of manually input new amounts' column in Excel table."
'@VariableDescription("Index of manually input last change dates' column in Excel table.")
Private Const m_manChangeDateColumn As Long = 12
Attribute m_manChangeDateColumn.VB_VarDescription = "Index of manually input last change dates' column in Excel table."

' Other constant
'@VariableDescription("How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations.")
Private Const m_diffThresholdPercent As Double = 0.01
Attribute m_diffThresholdPercent.VB_VarDescription = "How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations."

' 覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧覧� '


'@EntryPoint
'@Description("Deletes all entries who's automatic and manual last changed dates are before the checkup date.")
Public Sub DeleteUnchanged()
Attribute DeleteUnchanged.VB_Description = "Deletes all entries who's automatic and manual last changed dates are before the checkup date."
    Dim i As Long
    i = StartingRow
    Dim manWasChanged As Boolean
    Dim autoWasChanged As Boolean
    Do Until LenB(ActiveSheet.Cells(i, ItemColumn).Value) = 0
        manWasChanged = GetActCellValue(i, m_manChangeDateColumn) = m_checkupDate
        autoWasChanged = GetActCellValue(i, m_autoChangeDateColumn) = m_checkupDate
        If manWasChanged Or autoWasChanged Then
            i = i + 1
        Else
            ActiveSheet.Rows(i).Delete
        End If
    Loop
End Sub

'@EntryPoint
'@Description("Deletes all entries without difference in BB-date and (significant to a threshold) difference in automatic and manual new amount.")
Public Sub DeleteEquals()
Attribute DeleteEquals.VB_Description = "Deletes all entries without difference in BB-date and (significant to a threshold) difference in automatic and manual new amount."
    Dim i As Long
    i = StartingRow
    Dim bbDateMatch As Boolean
    Dim diff As Double
    Dim diffThreshold As Double
    Dim diffMatch As Boolean
    Do Until LenB(GetActCellValue(i, ItemColumn)) = 0
        bbDateMatch = CDate(GetActCellValue(i, m_autoBBDateColumn)) = CDate(GetActCellValue(i, m_manBBDateColumn))
        diff = Abs(CDbl(GetActCellValue(i, m_autoNewAmountColumn)) - CDbl(GetActCellValue(i, m_manNewAmountColumn)))
        diffThreshold = GetActCellValue(i, m_autoNewAmountColumn) * (m_diffThresholdPercent / 100)
        diffMatch = diff < diffThreshold
        If bbDateMatch And diffMatch Then
            ActiveSheet.Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
End Sub
