Attribute VB_Name = "DoesStringEndWithTest"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'Right$ is fastest.
'InStrB & StrReverse is a lot slower, InStr & StrReverse a little slower than that and Like again a little slower.
'Right is slowest by far, almost doubling Right$

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const IterationCountLabel As String = "Number of iterations: "
Attribute IterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const FirstMethodLabel As String = "Using InStr & StrReverse: "
Attribute FirstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const SecondMethodLabel As String = "Using InStrB & StrReverse: "
Attribute SecondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const ThirdMethodLabel As String = "Using Right =: "
Attribute ThirdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Label text for fourth testing method."
Private Const FourthMethodLabel As String = "Using Right$ =: "
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
    TestStr = "ExampleStringThatsLongerThanUsualToHaveBetterGroundWorkForTheDifferentApplicationsTest"
    Dim TestForStr As String
    TestForStr = "Test"
    Dim TestForStrRev As String
    TestForStrRev = "tseT"

    'Using InStr & StrReverse
    StarttimeOne = Timer
    For i = 1 To IterationCount
        TestBool = InStr(StrReverse(TestStr), TestForStrRev) = 1
    Next i
    EndtimeOne = Timer

    'Using InStrB & StrReverse
    StarttimeTwo = Timer
    For i = 1 To IterationCount
        TestBool = InStrB(StrReverse(TestStr), TestForStrRev) = 1
    Next i
    EndtimeTwo = Timer
    
    'Using Left =
    StarttimeThree = Timer
    For i = 1 To IterationCount
        '@Ignore UntypedFunctionUsage
        TestBool = Right(TestStr, Len(TestForStr)) = TestForStr
    Next i
    EndtimeThree = Timer
    
    'Using Left$ =
    StarttimeFour = Timer
    For i = 1 To IterationCount
        '@Ignore UntypedFunctionUsage
        TestBool = Right$(TestStr, Len(TestForStr)) = TestForStr
    Next i
    EndtimeFour = Timer

    'Using Like
    '@Ignore AssignmentNotUsed
    TestForStr = "*" & TestForStr
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
