Attribute VB_Name = "AssignEmptyStringTest"
Attribute VB_Description = "String emptiness testing."
'@Folder "Speed tests"
'@ModuleDescription "Empty string asssignment testing."
Option Explicit

'Result:
'Assigning vbNullString is much faster

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const IterationCountLabel As String = "Number of iterations: "
Attribute IterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const FirstMethodLabel As String = "Using vbNullString: "
Attribute FirstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const SecondMethodLabel As String = "Using """": "
Attribute SecondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Format of decimal number string output."
Private Const NumberFormat As String = "0.####"
Attribute NumberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription "Unit of measured time."
Private Const Unit As String = " s"
Attribute Unit.VB_VarDescription = "Unit of measured time."

'Count
'@VariableDescription "Amount of iterations to do for testing."
Private Const IterationCount As Long = 100000000
Attribute IterationCount.VB_VarDescription = "Amount of iterations to do for testing."


'@EntryPoint
'@Description "Tests string emptiness with different methods."
Public Sub TestAssignEmptyString()
Attribute TestAssignEmptyString.VB_Description = "Tests string emptiness with different methods."
    Dim i As Long
    Dim StarttimeOne As Double
    Dim EndtimeOne As Double
    Dim StarttimeTwo As Double
    Dim EndtimeTwo As Double
    Dim Msg As String

    'Test-variable
    '@Ignore VariableNotUsed
    Dim TestStr As String

    'Using '= vbNullString'
    StarttimeOne = Timer
    For i = 1 To IterationCount
        TestStr = vbNullString
    Next i
    EndtimeOne = Timer

    'Using '= ""'
    StarttimeTwo = Timer
    For i = 1 To IterationCount
        '@Ignore EmptyStringLiteral
        TestStr = ""
    Next i
    EndtimeTwo = Timer
    
    Msg = IterationCountLabel & "10^" & Log(IterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          FirstMethodLabel & Format$(EndtimeOne - StarttimeOne, NumberFormat) & Unit & vbNewLine & _
          SecondMethodLabel & Format$(EndtimeTwo - StarttimeTwo, NumberFormat) & Unit
    MsgBox Msg
End Sub
