Attribute VB_Name = "IsStringInString"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'InStr is fastest, closely followed by InStrB
'Using Like is much slower

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const IterationCountLabel As String = "Number of iterations: "
Attribute IterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const FirstMethodLabel As String = "Using InStr: "
Attribute FirstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const SecondMethodLabel As String = "Using InStrB: "
Attribute SecondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const ThirdMethodLabel As String = "Using Like: "
Attribute ThirdMethodLabel.VB_VarDescription = "Label text for third testing method."
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
'@Description "Tests if string contains other string with different methods."
Public Sub TestIsStringInString()
Attribute TestIsStringInString.VB_Description = "Tests if string contains other string with different methods."
    Dim i As Long
    Dim StarttimeOne As Double
    Dim EndtimeOne As Double
    Dim StarttimeTwo As Double
    Dim EndtimeTwo As Double
    Dim StarttimeThree As Double
    Dim EndtimeThree As Double
    Dim Msg As String
    '@Ignore VariableNotUsed
    Dim TestBool As Boolean

    'Test-variable
    Dim TestStr As String
    TestStr = "Example,Test"
    Dim TestForStr As String
    TestForStr = ","

    'Using InStr
    StarttimeOne = Timer
    For i = 1 To IterationCount
        TestBool = InStr(TestStr, TestForStr) > 0
    Next i
    EndtimeOne = Timer

    'Using InStrB
    StarttimeTwo = Timer
    For i = 1 To IterationCount
        TestBool = InStrB(TestStr, TestForStr) > 0
    Next i
    EndtimeTwo = Timer

    'Using Like
    '@Ignore AssignmentNotUsed
    TestForStr = "*" & TestForStr & "*"
    StarttimeThree = Timer
    For i = 1 To IterationCount
        TestBool = TestStr Like TestForStr
    Next i
    EndtimeThree = Timer
    
    
    Msg = IterationCountLabel & "10^" & Log(IterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          FirstMethodLabel & Format$(EndtimeOne - StarttimeOne, NumberFormat) & Unit & vbNewLine & _
          SecondMethodLabel & Format$(EndtimeTwo - StarttimeTwo, NumberFormat) & Unit & vbNewLine & _
          ThirdMethodLabel & Format$(EndtimeThree - StarttimeThree, NumberFormat) & Unit
    MsgBox Msg
End Sub
