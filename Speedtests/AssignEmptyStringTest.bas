Attribute VB_Name = "AssignEmptyStringTest"
Attribute VB_Description = "Module for testing string emptiness with different methods."
'@Folder "Speed tests"
'@ModuleDescription "Module for testing string emptiness with different methods."
Option Explicit

'Result:
'Using vbNullString is much faster

'String constants
Private Const IterationCountLabel As String = "Number of iterations: "
Private Const FirstMethodLabel As String = "Using vbNullString: "
Private Const SecondMethodLabel As String = "Using """": "
Private Const NumberFormat As String = "0.####"
Private Const Unit As String = " s"

'Count
Private Const IterationCount As Long = 100000000


'@EntryPoint
'@Description "Method for testing string emptiness with different methods."
Public Sub TestAssignEmptyString()
Attribute TestAssignEmptyString.VB_Description = "Method for testing string emptiness with different methods."
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
