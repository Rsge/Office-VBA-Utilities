Attribute VB_Name = "DoesStringEndWithTest"
Attribute VB_Description = "String contains string testing."
'@Folder("Tests")
'@ModuleDescription("String contains string testing.")
Option Explicit

' Result:
' Right$ & Len is fastest.
' Both 'InStrBs, Right & Len's are a bit slower, both 'InStr, Right & Len's a little more.
' Both 'InStr & StrReverse's are slower still, closely followed by Like.
' Right & Len is slowest, almost doubling Right$ & Len.

' Runtime constants
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using InStr & StrReverse: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using InStrB & StrReverse: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Label text for third testing method.")
Private Const m_3rdMethodLabel As String = "Using Right & Len: "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription("Label text for fourth testing method.")
Private Const m_4thMethodLabel As String = "Using Right$ & Len: "
Attribute m_4thMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription("Label text for fifth testing method")
Private Const m_5thMethodLabel As String = "Using InStr, Right & Len: "
Attribute m_5thMethodLabel.VB_VarDescription = "Label text for fifth testing method"
'@VariableDescription("Label text for sixth testing method")
Private Const m_6thMethodLabel As String = "Using InStr, Right$ & Len: "
Attribute m_6thMethodLabel.VB_VarDescription = "Label text for sixth testing method"
'@VariableDescription("Label text for seventh testing method")
Private Const m_7thMethodLabel As String = "Using InStrB, Right & Len: "
Attribute m_7thMethodLabel.VB_VarDescription = "Label text for seventh testing method"
'@VariableDescription("Label text for eighth testing method")
Private Const m_8thMethodLabel As String = "Using InStrB, Right$ & Len: "
Attribute m_8thMethodLabel.VB_VarDescription = "Label text for eighth testing method"
'@VariableDescription("Label text for ninth testing method.")
Private Const m_9thMethodLabel As String = "Using Like: "
Attribute m_9thMethodLabel.VB_VarDescription = "Label text for ninth testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 9
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Tests if string starts with other string with different methods.")
Public Sub TestStringEndsWith()
Attribute TestStringEndsWith.VB_Description = "Tests if string starts with other string with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel, _
                    m_4thMethodLabel, m_5thMethodLabel, m_6thMethodLabel, _
                    m_7thMethodLabel, m_8thMethodLabel, m_9thMethodLabel)
    ' Define test variables.
    Dim testStr As String
    testStr = "ExampleStringThatsLongerThanUsualToHaveBetterGroundWorkForTheDifferentApplicationsTest"
    Dim testForStr As String
    testForStr = "Test"
    Dim testForStrRev As String
    testForStrRev = "tseT"
    '@Ignore VariableNotUsed
    Dim testBool As Boolean
    
    ' Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ IterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String
    
    ' Test InStr & StrReverse.
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testBool = InStr(StrReverse(testStr), testForStrRev) = 1
    Next
    endTimes(0) = Timer

    ' Test InStrB & StrReverse.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(StrReverse(testStr), testForStrRev) = 1
    Next
    endTimes(1) = Timer
    
    ' Test Right & Len.
    startTimes(2) = Timer
    For i = 0 To iterationCount
        '@Ignore UntypedFunctionUsage
        testBool = Right(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(2) = Timer
    
    ' Test Right$ & Len.
    startTimes(3) = Timer
    For i = 0 To iterationCount
        testBool = Right$(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(3) = Timer

    ' Test InStr, Right & Len.
    startTimes(4) = Timer
    For i = 0 To iterationCount
        testBool = InStr(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(4) = Timer
    
    ' Test InStr, Right$ & Len.
    startTimes(5) = Timer
    For i = 0 To iterationCount
        testBool = InStr(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(5) = Timer
    
    ' Test InStrB, Right & Len.
    startTimes(6) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(6) = Timer
    
    ' Test InStrB, Right$ & Len.
    startTimes(7) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(Right$(testStr, Len(testForStr)), testForStr) > 0
    Next
    endTimes(7) = Timer
    
    ' Test Like.
    '@Ignore AssignmentNotUsed
    testForStr = "*" & testForStr
    startTimes(8) = Timer
    For i = 0 To iterationCount
        testBool = testStr Like testForStr
    Next
    endTimes(8) = Timer
    
    ' Output results.
    msg = IterationCountLabel & "10^" & IterationsExponent & vbNewLine & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), NumberFormat) & Unit
    Next
    MsgBox msg
End Sub
