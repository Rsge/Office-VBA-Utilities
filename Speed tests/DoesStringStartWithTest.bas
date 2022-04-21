Attribute VB_Name = "DoesStringStartWithTest"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'InStrB ist fastest, InStr slightly slower and Like a little slower again.
'Left$ & Len is much slower and Left & Len again much slower, almost doubling Left$ & Len and more than quintupling InStrB.

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
Private Const ThirdMethodLabel As String = "Using Left & Len =: "
Attribute ThirdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Label text for fourth testing method."
Private Const FourthMethodLabel As String = "Using Left$ & Len =: "
Attribute FourthMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription "Label text for fith testing method."
Private Const FithMethodLabel As String = "Using Like: "
Attribute FithMethodLabel.VB_VarDescription = "Label text for fith testing method."
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
'@Description "Tests if string starts with other string with different methods."
Public Sub TestStringStartsWith()
Attribute TestStringStartsWith.VB_Description = "Tests if string starts with other string with different methods."
    Dim i As Long
    Dim StarttimeOne As Double
    Dim EndtimeOne As Double
    Dim StarttimeTwo As Double
    Dim EndtimeTwo As Double
    Dim StarttimeThree As Double
    Dim EndtimeThree As Double
    Dim StarttimeFour As Double
    Dim EndtimeFour As Double
    Dim StarttimeFive As Double
    Dim EndtimeFive As Double
    Dim Msg As String
    '@Ignore VariableNotUsed
    Dim TestBool As Boolean

    'Test-variable
    Dim TestStr As String
    TestStr = "TestExample"
    Dim TestForStr As String
    TestForStr = "Test"

    'Using InStr
    StarttimeOne = Timer
    For i = 1 To IterationCount
        TestBool = InStr(TestStr, TestForStr) = 1
    Next i
    EndtimeOne = Timer

    'Using InStrB
    StarttimeTwo = Timer
    For i = 1 To IterationCount
        TestBool = InStrB(TestStr, TestForStr) = 1
    Next i
    EndtimeTwo = Timer
    
    'Using Left & Len =
    StarttimeThree = Timer
    For i = 1 To IterationCount
        '@Ignore UntypedFunctionUsage
        TestBool = Left(TestStr, Len(TestForStr)) = TestForStr
    Next i
    EndtimeThree = Timer
    
    'Using Left$ & Len =
    StarttimeFour = Timer
    For i = 1 To IterationCount
        TestBool = Left$(TestStr, Len(TestForStr)) = TestForStr
    Next i
    EndtimeFour = Timer

    'Using Like
    '@Ignore AssignmentNotUsed
    TestForStr = TestForStr & "*"
    StarttimeFive = Timer
    For i = 1 To IterationCount
        TestBool = TestStr Like TestForStr
    Next i
    EndtimeFive = Timer
    
    
    Msg = IterationCountLabel & "10^" & Log(IterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          FirstMethodLabel & Format$(EndtimeOne - StarttimeOne, NumberFormat) & Unit & vbNewLine & _
          SecondMethodLabel & Format$(EndtimeTwo - StarttimeTwo, NumberFormat) & Unit & vbNewLine & _
          ThirdMethodLabel & Format$(EndtimeThree - StarttimeThree, NumberFormat) & Unit & vbNewLine & _
          FourthMethodLabel & Format$(EndtimeFour - StarttimeFour, NumberFormat) & Unit & vbNewLine & _
          FithMethodLabel & Format$(EndtimeFive - StarttimeFive, NumberFormat) & Unit
    MsgBox Msg
End Sub
