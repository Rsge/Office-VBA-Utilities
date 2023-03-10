Attribute VB_Name = "AssignEmptyStringTest"
Attribute VB_Description = "Tests empty string assignment methods."
'@Folder("SpeedTests.Tests")
'@ModuleDescription("Tests empty string assignment methods.")
Option Explicit

' Result:
' Assigning vbNullString is much faster.

' Runtime constants
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using vbNullString: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using """": "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 2
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Tests string emptiness with different methods.")
Public Sub TestAssignEmptyString()
Attribute TestAssignEmptyString.VB_Description = "Tests string emptiness with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel)
    'Define test variables.
    '@Ignore VariableNotUsed
    Dim testStr As String
    
    ' Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ IterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String

    ' Test '= vbNullString'.
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testStr = vbNullString
    Next
    endTimes(0) = Timer

    ' Test '= ""'.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        '@Ignore EmptyStringLiteral
        testStr = ""
    Next
    endTimes(1) = Timer
    
    ' Output results.
    msg = IterationCountLabel & "10^" & IterationsExponent & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), NumberFormat) & Unit
    Next
    MsgBox msg
End Sub
