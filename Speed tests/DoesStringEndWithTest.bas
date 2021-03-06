Attribute VB_Name = "DoesStringEndWithTest"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'Right$ & Len is fastest.
'Both 'InStrBs, Right & Len's are a bit slower, both 'InStr, Right & Len's a little more.
'Both 'InStr & StrReverse's are slower still, closely followed by Like.
'Right & Len is slowest, almost doubling Right$ & Len.

'Constants to change
'@VariableDescription "Label text for first testing method."
Private Const m_1stMethodLabel As String = "Using InStr & StrReverse: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const m_2ndMethodLabel As String = "Using InStrB & StrReverse: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const m_3rdMethodLabel As String = "Using Right & Len: "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Label text for fourth testing method."
Private Const m_4thMethodLabel As String = "Using Right$ & Len: "
Attribute m_4thMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription "Label text for fifth testing method
Private Const m_5thMethodLabel As String = "Using InStr, Right & Len: "
'@VariableDescription "Label text for sixth testing method
Private Const m_6thMethodLabel As String = "Using InStr, Right$ & Len: "
'@VariableDescription "Label text for seventh testing method
Private Const m_7thMethodLabel As String = "Using InStrB, Right & Len: "
'@VariableDescription "Label text for eighth testing method
Private Const m_8thMethodLabel As String = "Using InStrB, Right$ & Len: "
'@VariableDescription "Label text for ninth testing method."
Private Const m_9thMethodLabel As String = "Using Like: "
Attribute m_9thMethodLabel.VB_VarDescription = "Label text for ninth testing method."
'@VariableDescription "Number of methods to test."
Private Const m_methodsCount As Long = 9
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."
'@VariableDescription "Exponent of 10 for amount of iterations to do in testing."
Private Const m_iterationsExponent As Long = 8
Attribute m_iterationsExponent.VB_VarDescription = "Exponent of 10 for amount of iterations to do in testing."

'Constants
'@VariableDescription "Label text for iteration count output."
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Format of decimal number string output."
Private Const m_numberFormat As String = "0.0###"
Attribute m_numberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription "Unit of measured time."
Private Const m_unit As String = " s"
Attribute m_unit.VB_VarDescription = "Unit of measured time."


'@EntryPoint
'@Description "Tests if string starts with other string with different methods."
Public Sub TestStringEndsWith()
Attribute TestStringEndsWith.VB_Description = "Tests if string starts with other string with different methods."
    'Inserting method labels in array
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel, _
                    m_4thMethodLabel, m_5thMethodLabel, m_6thMethodLabel, _
                    m_7thMethodLabel, m_8thMethodLabel, m_9thMethodLabel)
    'Defining test variables
    Dim testStr As String
    testStr = "ExampleStringThatsLongerThanUsualToHaveBetterGroundWorkForTheDifferentApplicationsTest"
    Dim testForStr As String
    testForStr = "Test"
    Dim testForStrRev As String
    testForStrRev = "tseT"
    '@Ignore VariableNotUsed
    Dim testBool As Boolean
    
    'Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ m_iterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String
    
    'Using InStr & StrReverse
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testBool = InStr(StrReverse(testStr), testForStrRev) = 1
    Next
    endTimes(0) = Timer

    'Using InStrB & StrReverse
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(StrReverse(testStr), testForStrRev) = 1
    Next
    endTimes(1) = Timer
    
    'Using Right & Len
    startTimes(2) = Timer
    For i = 0 To iterationCount
        '@Ignore UntypedFunctionUsage
        testBool = Right(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(2) = Timer
    
    'Using Right$ & Len
    startTimes(3) = Timer
    For i = 0 To iterationCount
        testBool = Right$(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(3) = Timer

    'Using InStr, Right & Len
    startTimes(4) = Timer
    For i = 0 To iterationCount
        testBool = InStr(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(4) = Timer
    
    'Using InStr, Right$ & Len
    startTimes(5) = Timer
    For i = 0 To iterationCount
        testBool = InStr(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(5) = Timer
    
    'Using InStrB, Right & Len
    startTimes(6) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(6) = Timer
    
    'Using InStrB, Right$ & Len
    startTimes(7) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(7) = Timer
    
    'Using Like
    '@Ignore AssignmentNotUsed
    testForStr = "*" & testForStr
    startTimes(8) = Timer
    For i = 0 To iterationCount
        testBool = testStr Like testForStr
    Next
    endTimes(8) = Timer
    
    'Output results
    msg = m_iterationCountLabel & "10^" & m_iterationsExponent & vbNewLine & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), m_numberFormat) & m_unit
    Next
    MsgBox msg
End Sub
