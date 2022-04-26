Attribute VB_Name = "IsStringEmptyTest"
Attribute VB_Description = "String emptiness testing."
'@Folder "Speed tests"
'@ModuleDescription "String emptiness testing."
Option Explicit

'Result:
'LenB is fastest, closely followed by Len
'Using '= ""' is much slower, using '= vbNullString' is slowest by a bit

'String constants
'@VariableDescription "Label text for iteration count output."
Private Const m_iterationCountLabel As String = "Number of iterations: "
Attribute m_iterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription "Label text for first testing method."
Private Const m_firstMethodLabel As String = "Using Len: "
Attribute m_firstMethodLabel.VB_VarDescription = "Label text for first testing method."
'@VariableDescription "Label text for second testing method."
Private Const m_secondMethodLabel As String = "Using LenB: "
Attribute m_secondMethodLabel.VB_VarDescription = "Label text for second testing method."
'@VariableDescription "Label text for third testing method."
Private Const m_thirdMethodLabel As String = "Using '= vbNullString': "
Attribute m_thirdMethodLabel.VB_VarDescription = "Label text for third testing method."
'@VariableDescription "Label text for fourth testing method."
Private Const m_fourthMethodLabel As String = "Using '= """"': "
Attribute m_fourthMethodLabel.VB_VarDescription = "Label text for fourth testing method."
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
Public Sub TestIsStringEmpty()
Attribute TestIsStringEmpty.VB_Description = "Tests string emptiness with different methods."
    Dim i As Long
    Dim starttimeOne As Double
    Dim endtimeOne As Double
    Dim starttimeTwo As Double
    Dim endtimeTwo As Double
    Dim starttimeThree As Double
    Dim endtimeThree As Double
    Dim starttimeFour As Double
    Dim endtimeFour As Double
    Dim msg As String
    '@Ignore VariableNotUsed
    Dim testBool As Boolean

    'Test-variable
    Dim testStr As String
    testStr = vbNullString

    'Using Len
    starttimeOne = Timer
    For i = 1 To m_iterationCount
        testBool = Len(testStr) > 0
    Next i
    endtimeOne = Timer

    'Using LenB
    starttimeTwo = Timer
    For i = 1 To m_iterationCount
        testBool = LenB(testStr) > 0
    Next i
    endtimeTwo = Timer

    'Using '= vbNullString'
    starttimeThree = Timer
    For i = 1 To m_iterationCount
        testBool = testStr = vbNullString
    Next i
    endtimeThree = Timer
    
    'Using '= ""'
    starttimeFour = Timer
    For i = 1 To m_iterationCount
        '@Ignore EmptyStringLiteral
        testBool = testStr = ""
    Next i
    endtimeFour = Timer
    
    msg = m_iterationCountLabel & "10^" & Log(m_iterationCount) / Log(10) & vbNewLine & _
          vbNewLine & _
          m_firstMethodLabel & Format$(endtimeOne - starttimeOne, m_numberFormat) & m_unit & vbNewLine & _
          m_secondMethodLabel & Format$(endtimeTwo - starttimeTwo, m_numberFormat) & m_unit & vbNewLine & _
          m_thirdMethodLabel & Format$(endtimeThree - starttimeThree, m_numberFormat) & m_unit & vbNewLine & _
          m_fourthMethodLabel & Format$(endtimeFour - starttimeFour, m_numberFormat) & m_unit
    MsgBox msg
End Sub
