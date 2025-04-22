Attribute VB_Name = "CompareNumberTest"
Attribute VB_Description = "Template for speed tests."
'@Folder("SpeedTests.Tests")
'@ModuleDescription("Template for speed tests.")
Option Explicit

' Result:
' Using = is fastest.
' Using <> vs. > is nearly indistinguishable, but > seems marginally faster sometimes.

' Runtime constants
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using =: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using <>: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Label text for third testing method.")
Private Const m_3rdMethodLabel As String = "Using >: "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 3
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Tests something with different methods.")
Public Sub TestSomething()
Attribute TestSomething.VB_Description = "Tests something with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel)
    ' Define test variables.
    '@Ignore VariableNotUsed
    Dim testBool As Boolean

    ' Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ IterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String
    
    ' Test =.
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testBool = 10 = 0
    Next
    endTimes(0) = Timer

    ' Test <>.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = 10 <> 0
    Next
    endTimes(1) = Timer
    
    ' Test >.
    startTimes(2) = Timer
    For i = 0 To iterationCount
        testBool = 10 > 0
    Next
    endTimes(2) = Timer
    
    ' Output results.
    msg = IterationCountLabel & "10^" & IterationsExponent & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), NumberFormat) & Unit
    Next
    MsgBox msg
End Sub
