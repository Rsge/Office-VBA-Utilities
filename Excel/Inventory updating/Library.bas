Attribute VB_Name = "Library"
Attribute VB_Description = "General, often used methods."
'@Folder "Inventory updating"
'@ModuleDescription "General, often used methods."
Option Explicit

'@VariableDescription "Warning if no save location specified."
Private Const m_noSaveLocationWarning As String = "No save location configured or" & vbNewLine & _
                                                  "save location does not exist." & vbNewLine & _
                                                  "Please specify a save location."
Attribute m_noSaveLocationWarning.VB_VarDescription = "Warning if no save location specified."
'@VariableDescription "Warning if no files found at save location."
Private Const m_noFilesWarning As String = "No files found at specified location." & vbNewLine & _
                                           "Please specify a different folder or abort and add files first."
Attribute m_noFilesWarning.VB_VarDescription = "Warning if no files found at save location."


''@Description "Replaces occurences of {} in a string with the specified replacements."
'Public Function FormatString(ByVal str As String, ParamArray replacements() As Variant) As String
'    Dim strArray() As String
'    strArray = Split(str, "{}")
'    Dim out As Object
'    Set out = CreateObject("System.Collections.ArrayList")
'    Dim i As Long
'    i = 0
'    Dim replacement As Variant
'    For Each replacement In replacements
'        out.Add strArray(i)
'        out.Add replacement
'        i = i + 1
'    Next
'    out.Add strArray(i)
'    FormatString = Join(out.ToArray, vbNullString)
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
    Dim fileNumber As Long
    'Using first unused file number
    fileNumber = FreeFile
    Dim pointer As Long
    'String of fixed length 1
    Dim char As String * 1
    Dim currentLineNumber As Long
    currentLineNumber = 0
    Dim lastLines() As String
    ReDim lastLines(0 To lineCount - 1)

    'Opening file
    Open filePath For Binary As fileNumber
    'Setting pointer to last position in file
    pointer = LOF(fileNumber)
    Do
        'Reading char at position "Pointer" into "Char"
        Get fileNumber, pointer, char
        If char = vbCr Then
            'Simply skipping CRs for Linux compat
            pointer = pointer - 1
        ElseIf char = vbLf Then
            'Reading Count last lines of file
            If currentLineNumber < lineCount - 1 Then
                currentLineNumber = currentLineNumber + 1
                pointer = pointer - 1
            Else
                Exit Do
            End If
        Else
            pointer = pointer - 1
            'Adding char to result String
            lastLines(currentLineNumber) = char & lastLines(currentLineNumber)
        End If
    Loop
    Close fileNumber
    
    GetLastLine = lastLines
End Function

'@Description "Gets save location of data files and determines if files are available."
Public Function GetDataFilePath(ByVal pathCell As Range) As String
Attribute GetDataFilePath.VB_Description = "Gets save location of data files and determines if files are available."
    'Variables
    Dim path As String
    path = pathCell.Value
    Dim noFilesRepeat As Boolean

    Do
        If LenB(path) = 0 Or LenB(Dir(path, vbDirectory)) = 0 Then
            'Defining path
            If Not noFilesRepeat Then
                'MsgBox to cancel folder dialog
                If MsgBox(m_noSaveLocationWarning, vbOKCancel) = vbCancel Then Exit Function
                'Opening folder dialog
                Dim FolderDialog As FileDialog
                Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
                If FolderDialog.Show = 0 Then Exit Function
                'Getting path
                path = FolderDialog.SelectedItems.Item(1)
                pathCell.Value = path
            End If

            'Checking file existence
            If Len(Dir(path & Application.PathSeparator & Ext)) <> 0 Then
                noFilesRepeat = False
            Else
                If MsgBox(m_noFilesWarning, vbOKCancel) = vbCancel Then Exit Function
                path = vbNullString
                noFilesRepeat = True
            End If
        End If
    Loop While noFilesRepeat

    GetDataFilePath = path
End Function
