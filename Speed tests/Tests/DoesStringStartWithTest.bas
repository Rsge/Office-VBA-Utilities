Attribute VB_Name = "DoesStringStartWithTest"
Attribute VB_Description = "Tests string starting with string checking methods."
'@Folder("SpeedTests.Tests")
'@ModuleDescription("Tests string starting with string checking methods.")
Option Explicit

' Result:
' InStrB is fastest.
' InStr is slightly slower and Like a little slower again.
' Left$ & Len is much slower and Left & Len again much slower, almost doubling Left$ & Len and more than quintupling InStrB.

' Runtime constants
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using InStr: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using InStrB: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Label text for third testing method.")
Private Const m_3rdMethodLabel As String = "Using Left & Len: "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription("Label text for fourth testing method.")
Private Const m_4thMethodLabel As String = "Using Left$ & Len: "
Attribute m_4thMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription("Label text for fith testing method.")
Private Const m_5thMethodLabel As String = "Using Like: "
Attribute m_5thMethodLabel.VB_VarDescription = "Label text for fith testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 5
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Tests if string starts with other string with different methods.")
Public Sub TestStringStartsWith()
Attribute TestStringStartsWith.VB_Description = "Tests if string starts with other string with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel, _
                    m_4thMethodLabel, m_5thMethodLabel)
    ' Define test variables.
    Const testStr As String = "TestExample"
    Dim testForStr As String
    testForStr = "Test"
    '@Ignore VariableNotUsed
    Dim testBool As Boolean
    
    ' Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ IterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String

    ' Test InStr.
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testBool = InStr(testStr, testForStr) = 1
    Next
    endTimes(0) = Timer

    ' Test InStrB.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(testStr, testForStr) = 1
    Next
    endTimes(1) = Timer
    
    ' Test Left & Len =.
    startTimes(2) = Timer
    For i = 0 To iterationCount
        '@Ignore UntypedFunctionUsage
        testBool = Left(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(2) = Timer
    
    ' Test Left$ & Len =.
    startTimes(3) = Timer
    For i = 0 To iterationCount
        testBool = Left$(testStr, Len(testForStr)) = testForStr
    Next
    endTimes(3) = Timer

    ' Test Like.
    '@Ignore AssignmentNotUsed
    testForStr = testForStr & "*"
    startTimes(4) = Timer
    For i = 0 To iterationCount
        testBool = testStr Like testForStr
    Next
    endTimes(4) = Timer
    
    
    ' Output results.
    msg = IterationCountLabel & "10^" & IterationsExponent & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), NumberFormat) & Unit
    Next
    MsgBox msg
End Sub
