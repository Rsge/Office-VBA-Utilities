Attribute VB_Name = "IsStringEmptyTest"
Attribute VB_Description = "Module for testing string emptiness with different methods."
'@IgnoreModule InvalidAnnotation
'@Folder "Speed tests"
'@ModuleDescription "Module for testing string emptiness with different methods."
Option Explicit

'Result:
'LenB is fastest, closely followed by Len
'Using "" is much slower, using vbNullString is slowest by a bit

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const IterationCountLabel As String = "Number of iterations: "
'@VariableDescription "Label text for first testing method."
Private Const FirstMethodLabel As String = "Using Len: "
'@VariableDescription "Label text for second testing method."
Private Const SecondMethodLabel As String = "Using LenB: "
'@VariableDescription "Label text for third testing method."
Private Const ThirdMethodLabel As String = "Using '= vbNullString': "
'@VariableDescription "Label text for fourth testing method."
Private Const FourthMethodLabel As String = "Using '= """"': "
'@VariableDescription "Format of decimal number string output."
Private Const NumberFormat As String = "0.####"
'@VariableDescription "Unit of measured time."
Private Const Unit As String = " s"

'Count
'@VariableDescription "Amount of iterations to do for testing."
Private Const IterationCount As Long = 100000000


'@EntryPoint
'@Description "Method for testing string emptiness with different methods."
Public Sub TestIsStringEmpty()
Attribute TestIsStringEmpty.VB_Description = "Method for testing string emptiness with different methods."
    Dim i As Long
    Dim StarttimeOne As Double
    Dim EndtimeOne As Double
    Dim StarttimeTwo As Double
    Dim EndtimeTwo As Double
    Dim StarttimeThree As Double
    Dim EndtimeThree As Double
    Dim StarttimeFour As Double
    Dim EndtimeFour As Double
    Dim Msg As String
    '@Ignore VariableNotUsed
    Dim TestBool As Boolean

    'Test-variable
    Dim TestStr As String
    TestStr = vbNullString

    'Using Len
    StarttimeOne = Timer
    For i = 1 To IterationCount
        TestBool = Len(TestStr) > 0
    Next i
    EndtimeOne = Timer

    'Using LenB
    StarttimeTwo = Timer
    For i = 1 To IterationCount
        TestBool = LenB(TestStr) > 0
    Next i
    EndtimeTwo = Timer

    'Using '= vbNullString'
    StarttimeThree = Timer
    For i = 1 To IterationCount
        TestBool = TestStr = vbNullString
    Next i
    EndtimeThree = Timer
    
    'Using '= ""'
    StarttimeFour = Timer
    For i = 1 To IterationCount
        '@Ignore EmptyStringLiteral
        TestBool = TestStr = ""
    Next i
    EndtimeFour = Timer
    
    Msg = IterationCountLabel & "10^" & Log(IterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          FirstMethodLabel & Format$(EndtimeOne - StarttimeOne, NumberFormat) & Unit & vbNewLine & _
          SecondMethodLabel & Format$(EndtimeTwo - StarttimeTwo, NumberFormat) & Unit & vbNewLine & _
          ThirdMethodLabel & Format$(EndtimeThree - StarttimeThree, NumberFormat) & Unit & vbNewLine & _
          FourthMethodLabel & Format$(EndtimeFour - StarttimeFour, NumberFormat) & Unit
    MsgBox Msg
End Sub
