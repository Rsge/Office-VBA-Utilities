Attribute VB_Name = "Testing"
'@Folder("Inventory updating")
Option Explicit

'Excel table column constants (1-based)
''@VariableDescription "Index of items' column in Excel table."
'Private Const GenItemColumn As Long = 1
''@VariableDescription "Index of amounts' units' column in Excel table."
'Private Const GenUnitColumn As Long = 2
'@VariableDescription "Index of automatically processed BB-dates' column in Excel table."
Private Const AutoBBDateColumn As Long = 3
Attribute AutoBBDateColumn.VB_VarDescription = "Index of automatically processed BB-dates' column in Excel table."
''@VariableDescription "Index of automatically processed old amounts' column in Excel table."
'Private Const AutoOldAmountColumn As Long = 4
''@VariableDescription "Index of automatically processed amounts' differences' column in Excel table."
'Private Const AutoDiffAmountColumn As Long = 5
'@VariableDescription "Index of automatically processed new amounts' column in Excel table."
Private Const AutoNewAmountColumn As Long = 6
Attribute AutoNewAmountColumn.VB_VarDescription = "Index of automatically processed new amounts' column in Excel table."
'@VariableDescription "Index of automatically processed last change dates' column in Excel table."
Private Const AutoChangeDateColumn As Long = 7
Attribute AutoChangeDateColumn.VB_VarDescription = "Index of automatically processed last change dates' column in Excel table."
'@VariableDescription "Index of manually input BB-dates' column in Excel table."
Private Const ManBBDateColumn As Long = 8
Attribute ManBBDateColumn.VB_VarDescription = "Index of manually input BB-dates' column in Excel table."
''@VariableDescription "Index of manually input old amounts' column in Excel table."
'Private Const ManOldAmountColumn As Long = 9
''@VariableDescription "Index of manually input amounts' differences' column in Excel table."
'Private Const ManDiffAmountColumn As Long = 10
'@VariableDescription "Index of manually input new amounts' column in Excel table."
Private Const ManNewAmountColumn As Long = 11
Attribute ManNewAmountColumn.VB_VarDescription = "Index of manually input new amounts' column in Excel table."
'@VariableDescription "Index of manually input last change dates' column in Excel table."
Private Const ManChangeDateColumn As Long = 12
Attribute ManChangeDateColumn.VB_VarDescription = "Index of manually input last change dates' column in Excel table."

'@VariableDescription "Date the last import to account for has been done."
Private Const CheckupDate As String = "01.01.2000"
Attribute CheckupDate.VB_VarDescription = "Date the last import to account for has been done."
'@VariableDescription "How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations."
Private Const DiffThresholdPercent As Double = 0.01
Attribute DiffThresholdPercent.VB_VarDescription = "How much manual and automatic new value are allowed to differ in percent to account for imprecise floating point calculations."

'@EntryPoint
'@Description "Deletes all entries who's automatic and manual last changed dates are before the checkup date."
Public Sub DeleteUnchanged()
Attribute DeleteUnchanged.VB_Description = "Deletes all entries who's automatic and manual last changed dates are before the checkup date."
    Dim i As Long
    i = StartingRow
    Do Until LenB(ActiveSheet.Cells(i, ItemColumn).Value) = 0
        Dim ManWasChanged As Boolean
        ManWasChanged = ActiveSheet.Cells(i, ManChangeDateColumn).Value = CheckupDate
        Dim AutoWasChanged As Boolean
        AutoWasChanged = ActiveSheet.Cells(i, AutoChangeDateColumn).Value = CheckupDate
        If ManWasChanged Or AutoWasChanged Then
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
        Dim BBDateMatch As Boolean
        BBDateMatch = CDate(ActiveSheet.Cells(i, AutoBBDateColumn).Value) = CDate(ActiveSheet.Cells(i, ManBBDateColumn).Value)
        Dim Diff As Double
        Diff = Abs(CDbl(ActiveSheet.Cells(i, AutoNewAmountColumn).Value) - CDbl(ActiveSheet.Cells(i, ManNewAmountColumn).Value))
        Dim DiffThreshold As Double
        DiffThreshold = ActiveSheet.Cells(i, AutoNewAmountColumn).Value * (DiffThresholdPercent / 100)
        Dim DiffMatch As Boolean
        DiffMatch = Diff < DiffThreshold
        If BBDateMatch And DiffMatch Then
            ActiveSheet.Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
End Sub
