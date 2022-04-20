Attribute VB_Name = "Library"
Attribute VB_Description = "General, often used methods."
'@Folder "Inventory updating"
'@ModuleDescription "General, often used methods."
Option Explicit

'@VariableDescription "Warning if no save location specified."
Private Const NoSaveLocationWarning As String = "No save location configured or" & vbNewLine & _
                                                "save location does not exist." & vbNewLine & _
                                                "Please specify a save location."
Attribute NoSaveLocationWarning.VB_VarDescription = "Warning if no save location specified."
'@VariableDescription "Warning if no files found at save location."
Private Const NoFilesWarning As String = "No files found at specified location." & vbNewLine & _
                                         "Please specify a different folder or abort and add files first."
Attribute NoFilesWarning.VB_VarDescription = "Warning if no files found at save location."


''@Description "Replaces occurences of {} in a string with the specified replacements."
'Public Function FormatString(ByVal str As String, ParamArray replacements() As Variant) As String
'    Dim StrArray() As String
'    StrArray = Split(str, "{}")
'    Dim Out As Object
'    Set Out = CreateObject("System.Collections.ArrayList")
'    Dim i As Long
'    i = 0
'    Dim Replacement As Variant
'    For Each Replacement In replacements
'        Out.Add StrArray(i)
'        Out.Add Replacement
'        i = i + 1
'    Next
'    Out.Add StrArray(i)
'    FormatString = Join(Out.ToArray, vbNullString)
'End Function

'@Description "Tests if a string starts with another string."
Public Function StartsWith(ByVal str As String, ByVal start As String) As Boolean
Attribute StartsWith.VB_Description = "Tests if a string starts with another string."
    StartsWith = InStrB(str, start) = 1
End Function

'@Description "Tests if a string ends with another string."
Public Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
Attribute EndsWith.VB_Description = "Tests if a string ends with another string."
    EndsWith = Right$(str, Len(ending)) = ending
End Function

'@Description "Tests if a string contains another string."
Public Function Contains(ByVal str As String, ByVal match As String) As Boolean
Attribute Contains.VB_Description = "Tests if a string contains another string."
    Contains = InStr(str, match) > 0
End Function

'@Description "Tests if a string is empty."
Public Function IsEmpty(ByVal str As String) As Boolean
Attribute IsEmpty.VB_Description = "Tests if a string is empty."
    IsEmpty = LenB(str) = 0
End Function

'@Description "Gets name of file without the extension."
Public Function GetFileNameWithoutExtension(ByVal fileObject As Object) As String
Attribute GetFileNameWithoutExtension.VB_Description = "Gets name of file without the extension."
    GetFileNameWithoutExtension = Split(fileObject.Name, ".")(0)
End Function

'@Description "Gets number of last lines of file at path (default 2)."
Public Function GetLastLine(ByVal filePath As String, Optional ByVal lineCount As Long = 1) As String()
Attribute GetLastLine.VB_Description = "Gets number of last lines of file at path (default 2)."
    Dim FileNumber As Long
    'Using first unused file number
    FileNumber = FreeFile
    Dim Pointer As Long
    'String of fixed length 1
    Dim Char As String * 1
    Dim CurrentLineNumber As Long
    CurrentLineNumber = 0
    Dim LastLines() As String
    ReDim LastLines(0 To lineCount - 1)

    'Opening file
    Open filePath For Binary As FileNumber
    'Setting pointer to last position in file
    Pointer = LOF(FileNumber)
    Do
        'Reading char at position "Pointer" into "Char"
        Get FileNumber, Pointer, Char
        If Char = vbCr Then
            'Simply skipping CRs for Linux compat
            Pointer = Pointer - 1
        ElseIf Char = vbLf Then
            'Reading Count last lines of file
            If CurrentLineNumber < lineCount - 1 Then
                CurrentLineNumber = CurrentLineNumber + 1
                Pointer = Pointer - 1
            Else
                Exit Do
            End If
        Else
            Pointer = Pointer - 1
            'Adding char to result String
            LastLines(CurrentLineNumber) = Char & LastLines(CurrentLineNumber)
        End If
    Loop
    Close FileNumber
    
    GetLastLine = LastLines
End Function

'@Description "Gets save location of data files and determines if files are available."
Public Function GetDataFilePath(ByVal pathCell As Range) As String
Attribute GetDataFilePath.VB_Description = "Gets save location of data files and determines if files are available."
    'Variables
    Dim Path As String
    Path = pathCell.Value
    Dim NoFilesRepeat As Boolean

    Do
        If LenB(Path) = 0 Or LenB(Dir(Path, vbDirectory)) = 0 Then
            'Defining path
            If Not NoFilesRepeat Then
                'MsgBox to cancel folder dialog
                If MsgBox(NoSaveLocationWarning, vbOKCancel) = vbCancel Then Exit Function
                'Opening folder dialog
                Dim FolderDialog As FileDialog
                Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
                If FolderDialog.Show = 0 Then Exit Function
                'Getting path
                Path = FolderDialog.SelectedItems.Item(1)
                pathCell.Value = Path
            End If

            'Checking file existence
            If Len(Dir(Path & Application.PathSeparator & Ext)) <> 0 Then
                NoFilesRepeat = False
            Else
                If MsgBox(NoFilesWarning, vbOKCancel) = vbCancel Then Exit Function
                Path = vbNullString
                NoFilesRepeat = True
            End If
        End If
    Loop While NoFilesRepeat

    GetDataFilePath = Path
End Function
