Attribute VB_Name = "DoesStringEndWithTest"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'Right$ & Len is fastest.
'InStrB & StrReverse is a lot slower, InStr & StrReverse a little slower than that and Like again a little slower.
'Right & Len is slowest by far, almost doubling Right$

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const m_firstMethodLabel As String = "Using InStr & StrReverse: "
Attribute m_firstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const m_secondMethodLabel As String = "Using InStrB & StrReverse: "
Attribute m_secondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const m_thirdMethodLabel As String = "Using Right & Len =: "
Attribute m_thirdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Label text for fourth testing method."
Private Const m_fourthMethodLabel As String = "Using Right$ & Len =: "
Attribute m_fourthMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription "Label text for fith testing method."
Private Const m_fithMethodLabel As String = "Using Like: "
Attribute m_fithMethodLabel.VB_VarDescription = "Label text for fith testing method."
'@VariableDescription "Format of decimal number string output."
Private Const m_numberFormat As String = "0.####"
Attribute m_numberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription "Unit of measured time."
Private Const m_unit As String = " s"
Attribute m_unit.VB_VarDescription = "Unit of measured time."

'Count
'@VariableDescription "Amount of iterations to do for testing."
Private Const IterationCount As Long = 100000000
Attribute IterationCount.VB_VarDescription = "Amount of iterations to do for testing."


'@EntryPoint
'@Description "Tests if string starts with other string with different methods."
Public Sub TestStringStartsWith()
Attribute TestStringStartsWith.VB_Description = "Tests if string starts with other string with different methods."
    Dim i As Long
    Dim starttimeOne As Double
    Dim endtimeOne As Double
    Dim starttimeTwo As Double
    Dim endtimeTwo As Double
    Dim starttimeThree As Double
    Dim endtimeThree As Double
    Dim starttimeFour As Double
    Dim endtimeFour As Double
    Dim starttimeFive As Double
    Dim endtimeFive As Double
    Dim msg As String
    '@Ignore VariableNotUsed
    Dim testBool As Boolean

    'Test-variable
    Dim testStr As String
    testStr = "ExampleStringThatsLongerThanUsualToHaveBetterGroundWorkForTheDifferentApplicationsTest"
    Dim testForStr As String
    testForStr = "Test"
    Dim testForStrRev As String
    testForStrRev = "tseT"

    'Using InStr & StrReverse
    starttimeOne = Timer
    For i = 1 To IterationCount
        testBool = InStr(StrReverse(testStr), testForStrRev) = 1
    Next i
    endtimeOne = Timer

    'Using InStrB & StrReverse
    starttimeTwo = Timer
    For i = 1 To IterationCount
        testBool = InStrB(StrReverse(testStr), testForStrRev) = 1
    Next i
    endtimeTwo = Timer
    
    'Using Right & Len =
    starttimeThree = Timer
    For i = 1 To IterationCount
        '@Ignore UntypedFunctionUsage
        testBool = Right(testStr, Len(testForStr)) = testForStr
    Next i
    endtimeThree = Timer
    
    'Using Right$ & Len =
    starttimeFour = Timer
    For i = 1 To IterationCount
        testBool = Right$(testStr, Len(testForStr)) = testForStr
    Next i
    endtimeFour = Timer

    'Using Like
    '@Ignore AssignmentNotUsed
    testForStr = "*" & testForStr
    starttimeFive = Timer
    For i = 1 To IterationCount
        testBool = testStr Like testForStr
    Next i
    endtimeFive = Timer
    
    
    msg = m_iterationCountLabel & "10^" & Log(IterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          m_firstMethodLabel & Format$(endtimeOne - starttimeOne, m_numberFormat) & m_unit & vbNewLine & _
          m_secondMethodLabel & Format$(endtimeTwo - starttimeTwo, m_numberFormat) & m_unit & vbNewLine & _
          m_thirdMethodLabel & Format$(endtimeThree - starttimeThree, m_numberFormat) & m_unit & vbNewLine & _
          m_fourthMethodLabel & Format$(endtimeFour - starttimeFour, m_numberFormat) & m_unit & vbNewLine & _
          m_fithMethodLabel & Format$(endtimeFive - starttimeFive, m_numberFormat) & m_unit
    MsgBox msg
End Sub
