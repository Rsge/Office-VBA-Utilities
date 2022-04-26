Attribute VB_Name = "AssignEmptyStringTest"
Attribute VB_Description = "Empty string asssignment testing."
'@Folder "Speed tests"
'@ModuleDescription "Empty string asssignment testing."
Option Explicit

'Result:
'Assigning vbNullString is much faster

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const m_firstMethodLabel As String = "Using vbNullString: "
Attribute m_firstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const m_secondMethodLabel As String = "Using """": "
Attribute m_secondMethodLabel.VB_VarDescription = "Label text for second testing method."
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
'@Description "Tests string emptiness with different methods."
Public Sub TestAssignEmptyString()
Attribute TestAssignEmptyString.VB_Description = "Tests string emptiness with different methods."
    Dim i As Long
    Dim starttimeOne As Double
    Dim endtimeOne As Double
    Dim starttimeTwo As Double
    Dim endtimeTwo As Double
    Dim msg As String

    'Test-variable
    '@Ignore VariableNotUsed
    Dim testStr As String

    'Using '= vbNullString'
    starttimeOne = Timer
    For i = 1 To m_iterationCount
        testStr = vbNullString
    Next i
    endtimeOne = Timer

    'Using '= ""'
    starttimeTwo = Timer
    For i = 1 To m_iterationCount
        '@Ignore EmptyStringLiteral
        testStr = ""
    Next i
    endtimeTwo = Timer
    
    msg = m_iterationCountLabel & "10^" & Log(m_iterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          m_firstMethodLabel & Format$(endtimeOne - starttimeOne, m_numberFormat) & m_unit & vbNewLine & _
          m_secondMethodLabel & Format$(endtimeTwo - starttimeTwo, m_numberFormat) & m_unit
    MsgBox msg
End Sub
