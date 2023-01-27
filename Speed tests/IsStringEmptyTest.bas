Attribute VB_Name = "IsStringEmptyTest"
Attribute VB_Description = "String emptiness testing."
'@Folder("Tests")
'@ModuleDescription("String emptiness testing.")
Option Explicit

' Result:
' LenB is fastest, closely followed by Len.
' Using '= ""' is much slower, using '= vbNullString' is slowest by a bit.

' Constants to change
'@VariableDescription("Label text for first testing method.")
Private Const m_1stMethodLabel As String = "Using Len: "
Attribute m_1stMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription("Label text for second testing method.")
Private Const m_2ndMethodLabel As String = "Using LenB: "
Attribute m_2ndMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription("Label text for third testing method.")
Private Const m_3rdMethodLabel As String = "Using '= vbNullString': "
Attribute m_3rdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription("Label text for fourth testing method.")
Private Const m_4thMethodLabel As String = "Using '= """"': "
Attribute m_4thMethodLabel.VB_VarDescription = "Label text for fourth testing method."
'@VariableDescription("Number of methods to test.")
Private Const m_methodsCount As Long = 4
Attribute m_methodsCount.VB_VarDescription = "Number of methods to test."
'@VariableDescription("Exponent of 10 for amount of iterations to do in testing.")
Private Const m_iterationsExponent As Long = 8
Attribute m_iterationsExponent.VB_VarDescription = "Exponent of 10 for amount of iterations to do in testing."

' Constants
'@VariableDescription("Label text for iteration count output.")
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription("Format of decimal number string output.")
Private Const m_numberFormat As String = "0.0###"
Attribute m_numberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription("Unit of measured time.")
Private Const m_unit As String = " s"
Attribute m_unit.VB_VarDescription = "Unit of measured time."


'@EntryPoint
'@Description("Tests string emptiness with different methods.")
Public Sub TestIsStringEmpty()
Attribute TestIsStringEmpty.VB_Description = "Tests string emptiness with different methods."
    ' Insert method labels in array.
    Dim methods As Variant
    methods = Array(m_1stMethodLabel, m_2ndMethodLabel, m_3rdMethodLabel, m_4thMethodLabel)
    ' Define test variables.
    '@Ignore VariableNotUsed
    Dim testBool As Boolean
    Dim testStr As String
    testStr = vbNullString

    ' Other variables & constants
    Dim i As Long
    Const iterationCount As Long = 10 ^ m_iterationsExponent
    Const methodsLength As Long = m_methodsCount - 1
    Dim startTimes(0 To methodsLength) As Double
    Dim endTimes(0 To methodsLength) As Double
    Dim msg As String

    ' Test Len.
    startTimes(0) = Timer
    For i = 0 To iterationCount
        testBool = Len(testStr) > 0
    Next
    endTimes(0) = Timer

    ' Test LenB.
    startTimes(1) = Timer
    For i = 0 To iterationCount
        testBool = LenB(testStr) > 0
    Next
    endTimes(1) = Timer

    ' Test '= vbNullString'.
    startTimes(2) = Timer
    For i = 0 To iterationCount
        testBool = testStr = vbNullString
    Next
    endTimes(2) = Timer
    
    ' Test '= ""'.
    startTimes(3) = Timer
    For i = 0 To iterationCount
        '@Ignore EmptyStringLiteral
        testBool = testStr = ""
    Next
    endTimes(3) = Timer
    
    ' Output results.
    msg = m_iterationCountLabel & "10^" & m_iterationsExponent & vbNewLine
    For i = 0 To methodsLength
        msg = msg & vbNewLine & methods(i) & Format$(endTimes(i) - startTimes(i), m_numberFormat) & m_unit
    Next
    MsgBox msg
End Sub
