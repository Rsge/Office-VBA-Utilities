Attribute VB_Name = "PasswordSearch"
Attribute VB_Description = "Brute-force an Office password. Not really that fast, but it works..."
'@Folder("PasswordFinding")
'@ModuleDescription("Brute-force an Office password. Not really that fast, but it works...")
Option Explicit

' # Reference #
'@VariableDescription("Also try the first letter of the word capitalized.")
Public Const Capitalize As Boolean = True
Attribute Capitalize.VB_VarDescription = "Also try the first letter of the word capitalized."
'@VariableDescription("Question about minimum possible password length.")
Public Const MinLengthQuestion As String = "What is the minimum length the password could be? (Min: 1)"
Attribute MinLengthQuestion.VB_VarDescription = "Question about minimum possible password length."
'@VariableDescription("Question about minimum possible password length.")
Public Const FirstTryQuestion As String = "What is the first ('lowest') string to try?"
Attribute FirstTryQuestion.VB_VarDescription = "Question about minimum possible password length."
'@Description("The symbols used in variations.")
Public Function SymbolArr() As Variant
Attribute SymbolArr.VB_Description = "The symbols used in variations."
    SymbolArr = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
End Function

' ————————————————————————————————————————————————————— '

' # Library #
'@Description("Tests if a string ends with another string.")
Public Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
Attribute EndsWith.VB_Description = "Tests if a string ends with another string."
    EndsWith = Right$(str, Len(ending)) = ending
End Function

'@Description("Tests if a string is empty.")
Public Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

'@Description("Finds the max of two longs.")
Public Function Max(ByVal x As Long, ByVal y As Long) As Long
Attribute Max.VB_Description = "Finds the max of two longs."
    Max = IIf(x > y, x, y)
End Function

' ————————————————————————————————————————————————————— '

'@Ignore ParameterCanBeByVal
'@Description("Opens a protected document given the application's docs, the application-dependent mode and a password.")
Private Function TryOpenDoc(ByVal mode As Long, ByRef docs As Object, ByVal filePath As String, ByVal passwd As String) As Object
Attribute TryOpenDoc.VB_Description = "Opens a protected document given the application's docs, the application-dependent mode and a password."
    Select Case mode
    Case 0 ' doc
        Set TryOpenDoc = docs.Open(filePath, PasswordDocument:=passwd)
    Case 1 ' xls
        Set TryOpenDoc = docs.Open(filePath, Password:=passwd)
    Case 2 ' ppt
        Set TryOpenDoc = docs.Open(filePath + "::" + passwd + "::")
    End Select
End Function

'@Description("Creates a 2-dim array of the cartesian product of each element of a given array with each main entry a new variant and the subarray making up that variant.")
Private Function CartesianProduct(ByRef arr() As Variant, ByVal k As Long) As Variant
Attribute CartesianProduct.VB_Description = "Creates a 2-dim array of the cartesian product of each element of a given array with each main entry a new variant and the subarray making up that variant."
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1 ' Len of array
    Dim vars As Long
    vars = n ^ k
    Dim result() As Variant
    ReDim result(0 To vars - 1, 0 To k - 1)
    Dim i As Long ' Variations
    Dim j As Long ' Indices
    Dim val As Long ' Value
    For i = 0 To vars - 1
        val = i
        For j = k - 1 To 0 Step -1
            result(i, j) = arr(val Mod n)
            val = val \ n
        Next
    Next
    CartesianProduct = result
End Function

' ————————————————————————————————————————————————————— '

'@EntryPoint
Public Sub FindPassword()
    ' Get symbols from Reference.
    Dim symbols() As Variant
    symbols = SymbolArr()
    ' Get file.
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Office", "*.doc?; *.xls?; *.ppt?", 1
        If .Show = -1 Then
            filePath = .SelectedItems.Item(1)
        End If
    End With
    ' Determine mode. (Which Office app?)
    Dim mode As Long
    Dim docs As Object
    If EndsWith(filePath, "doc") Or EndsWith(filePath, "docx") Or EndsWith(filePath, "docm") Then
        mode = 0
        Set docs = CreateObject("Word.Application").Documents
    ElseIf EndsWith(filePath, "xls") Or EndsWith(filePath, "xlsx") Or EndsWith(filePath, "xlsm") Then
        mode = 1
        Set docs = Workbooks
    ElseIf EndsWith(filePath, "ppt") Or EndsWith(filePath, "pptx") Or EndsWith(filePath, "pptm") Then
        mode = 2
        Set docs = CreateObject("PowerPoint.Application").Presentations
    Else
        Exit Sub
    End If
    ' Get minimum possible password length and first string to try.
    Dim minLen As Long
    Do While True
        Dim inputStr As String
        inputStr = InputBox(MinLengthQuestion, Default:=1)
        If IsEmpty(inputStr) Then Exit Sub
        On Error Resume Next
        minLen = CInt(inputStr)
        On Error GoTo 0
        If minLen > 0 Then Exit Do
    Loop
    Dim firstTry As String
    firstTry = InputBox(FirstTryQuestion)
    '@Ignore AssignmentNotUsed
    minLen = Max(minLen, Len(firstTry))
    ' Brute-force password.
    Dim k As Long
    k = minLen
    Dim found As Boolean
    found = False
    Dim passwd As String
    Dim doc As Object
    Do Until found
        ' Generate tries.
        Dim tries As Variant
        tries = CartesianProduct(symbols, k)
        Dim i As Long
        Dim j As Long
        For i = 0 To UBound(tries, 1)
            ' Generate password string.
            passwd = vbNullString
            For j = 0 To UBound(tries, 2)
                passwd = passwd + tries(i, j)
            Next
            ' Reduce lag with still high speed.
            If i Mod 50 = 0 Then
                DoEvents
            End If
            ' Skip entries before first try.
            If StrComp(passwd, firstTry, vbTextCompare) < 0 Then GoTo Continue
            ' Try password.
            Debug.Print "Trying: " + passwd
            On Error Resume Next
            Set doc = TryOpenDoc(mode, docs, filePath, passwd)
            If Capitalize Then
                If doc Is Nothing Then
                    ' Try the password capitalized.
                    passwd = StrConv(passwd, vbProperCase)
                    Debug.Print "Trying: " + passwd
                    Set doc = TryOpenDoc(mode, docs, filePath, passwd)
                End If
            End If
            On Error GoTo 0
            ' If the doc was opened, close it and exit the loops.
            If Not doc Is Nothing Then
                doc.Close
                found = True
                Exit For
            End If
Continue:
        Next
        k = k + 1
    Loop
    MsgBox "The password is: " + passwd
End Sub

' ————————————————————————————————————————————————————— '

'@EntryPoint
Public Sub TestProduct()
    Dim symbols() As Variant
    symbols = SymbolArr
    Dim k As Long
    k = 2 ' Number of times to repeat
    Dim tries As Variant
    tries = CartesianProduct(symbols, k)
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(tries, 1)
        For j = 0 To UBound(tries, 2)
            Debug.Print tries(i, j);
        Next
        Debug.Print
    Next
End Sub
