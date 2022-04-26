Attribute VB_Name = "IsStringInStringTest"
Attribute VB_Description = "String contains string testing."
'@Folder "Speed tests"
'@ModuleDescription "String contains string testing."
Option Explicit

'Result:
'InStr is fastest, closely followed by InStrB
'Using Like is much slower

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const m_firstMethodLabel As String = "Using InStr: "
Attribute m_firstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const m_secondMethodLabel As String = "Using InStrB: "
Attribute m_secondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const m_thirdMethodLabel As String = "Using Like: "
Attribute m_thirdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Format of decimal number string output."
Private Const m_numberFormat As String = "0.####"
Attribute m_numberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription "Unit of measured time."
Private Const m_unit As String = " s"
Attribute m_unit.VB_VarDescription = "Unit of measured time."

'Count
'@VariableDescription "Amount of iterations to do for testing."
Private Const m_iterationCount As Long = 100000000
Attribute m_iterationCount.VB_VarDescription = "Amount of iterations to do for testing."


'@EntryPoint
'@Description "Tests if string contains other string with different methods."
Public Sub TestIsStringInString()
Attribute TestIsStringInString.VB_Description = "Tests if string contains other string with different methods."
    Dim i As Long
    Dim starttimeOne As Double
    Dim endtimeOne As Double
    Dim starttimeTwo As Double
    Dim endtimeTwo As Double
    Dim starttimeThree As Double
    Dim endtimeThree As Double
    Dim msg As String
    '@Ignore VariableNotUsed
    Dim testBool As Boolean

    'Test-variable
    Dim testStr As String
    testStr = "Example,Test"
    Dim testForStr As String
    testForStr = ","

    'Using InStr
    starttimeOne = Timer
    For i = 1 To m_iterationCount
        testBool = InStr(testStr, testForStr) > 0
    Next i
    endtimeOne = Timer

    'Using InStrB
    starttimeTwo = Timer
    For i = 1 To m_iterationCount
        testBool = InStrB(testStr, testForStr) > 0
    Next i
    endtimeTwo = Timer

    'Using Like
    '@Ignore AssignmentNotUsed
    testForStr = "*" & testForStr & "*"
    starttimeThree = Timer
    For i = 1 To m_iterationCount
        testBool = testStr Like testForStr
    Next i
    endtimeThree = Timer
    
    
    msg = m_iterationCountLabel & "10^" & Log(m_iterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          m_firstMethodLabel & Format$(endtimeOne - starttimeOne, m_numberFormat) & m_unit & vbNewLine & _
          m_secondMethodLabel & Format$(endtimeTwo - starttimeTwo, m_numberFormat) & m_unit & vbNewLine & _
          m_thirdMethodLabel & Format$(endtimeThree - starttimeThree, m_numberFormat) & m_unit
    MsgBox msg
End Sub
