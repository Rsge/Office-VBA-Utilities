Attribute VB_Name = "IsStringInStringTest"
Attribute VB_Description = "String contains string testing."
'@Folder("Tests")
'@ModuleDescription("String contains string testing.")
Option Explicit

' Result:
' InStr is fastest, closely followed by InStrB.
' Using Like is much slower.

' Runtime constants
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using InStr: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using InStrB: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Label text for third testing method.")
Private Const m_3rdMethodLabel As String = "Using Like: "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 3
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Tests if string contains other string with different methods.")
Public Sub TestIsStringInString()
Attribute TestIsStringInString.VB_Description = "Tests if string contains other string with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel)
    ' Define test variables.
    '@Ignore VariableNotUsed
    Dim testBool As Boolean
    Dim testStr As String
    testStr = "Example,Test"
    Dim testForStr As String
    testForStr = ","
    
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
        testBool = InStr(testStr, testForStr) > 0
    Next
    endTimes(0) = Timer

    ' Test InStrB.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = InStrB(testStr, testForStr) > 0
    Next
    endTimes(1) = Timer

    ' Test Like.
    '@Ignore AssignmentNotUsed
    testForStr = "*" & testForStr & "*"
    startTimes(2) = Timer
    For i = 0 To iterationCount
        testBool = testStr Like testForStr
    Next
    endTimes(2) = Timer
    
    ' Output results.
    msg = IterationCountLabel & "10^" & IterationsExponent & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), NumberFormat) & Unit
    Next
    MsgBox msg
End Sub
