Attribute VB_Name = "FileRegexReplace"
Attribute VB_Description = "Handles regex text replacement in all files of a type in a starting folder and all it's subfolders."
'@Folder("FileEditing")
'@ModuleDescription("Handles regex text replacement in all files of a type in a starting folder and all it's subfolders.")
Option Explicit

'@VariableDescription("File extension of file type to process.")
Private Const m_fileExt As String = ".docx"
Attribute m_fileExt.VB_VarDescription = "File extension of file type to process."
'@VariableDescription("Starting folder of document search.")
Private Const m_basePath As String = "C:\Example"
Attribute m_basePath.VB_VarDescription = "Starting folder of document search."
'@VariableDescription("Delimiter of const string lists.")
Private Const m_sep As String = "`"
Attribute m_sep.VB_VarDescription = "Delimiter of const string lists."
'@VariableDescription("m_sep delimited list of regex patterns.")
Private Const m_regexPatternsStr As String = "(^\s$)`Example(\d)"
Attribute m_regexPatternsStr.VB_VarDescription = "m_sep delimited list of regex patterns."
'@VariableDescription("m_sep delimited list of replacements for regex matches.")
Private Const m_regexReplacementsStr As String = "`$1"
Attribute m_regexReplacementsStr.VB_VarDescription = "m_sep delimited list of replacements for regex matches."

' ————————————————————————————————————————————————————— '


'@Description("Tests if a string ends with another string.")
Public Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
Attribute EndsWith.VB_Description = "Tests if a string ends with another string."
    EndsWith = Right$(str, Len(ending)) = ending
End Function

'@Description("Makes changes in a file via regex replace.")
Private Sub ChangeStuff(ByVal filePath As String)
Attribute ChangeStuff.VB_Description = "Makes changes in a file via regex replace."
    ' Open doc.
    Dim doc As Document
    Set doc = Documents.Open(filePath)
    Dim changed As Boolean
    changed = False
    ' Set regex.
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim regexPatterns() As String
    regexPatterns = Split(m_regexPatternsStr, m_sep)
    Dim regexReplacements() As String
    regexReplacements = Split(m_regexReplacementsStr, m_sep)
    ' Replace in file.
    Dim par As Range
    Dim replaced As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    For i = 1 To doc.Paragraphs.Count
        Set par = doc.Paragraphs.Item(i).Range
        For j = LBound(regexPatterns) To UBound(regexPatterns)
            regex.Pattern = regexPatterns(j)
            replaced = regex.Replace(par.Text, regexReplacements(j))
            If replaced <> par.Text Then
                par.Text = replaced
                changed = True
                If j = LBound(regexPatterns) Then
                    k = 0
                    Do While doc.Paragraphs.Item(i - k).Range.Text = vbCr
                        doc.Paragraphs.Item(i - k).Range.Delete
                        k = k + 1
                    Loop
                End If
            End If
        Next
    Next
    ' Save and close file.
    If changed Then
        doc.Save
        changed = False
    End If
    doc.Close wdDoNotSaveChanges
End Sub

'@Description("Loops through all subfolders of base folder and through all their files.")
Private Sub LoopAllSubFolders(ByVal FSOFolder As Object)
Attribute LoopAllSubFolders.VB_Description = "Loops through all subfolders of base folder and through all their files."
    Dim FSOSubFolder As Object
    Dim FSOFile As Object
    ' Recurse into subfolders.
    For Each FSOSubFolder In FSOFolder.subfolders
        LoopAllSubFolders FSOSubFolder
    Next
    ' For each file of a specified type in a subfolder, change stuff.
    For Each FSOFile In FSOFolder.Files
        If EndsWith(FSOFile.Name, m_fileExt) Then
            ChangeStuff FSOFile.Path
            DoEvents
        End If
    Next
End Sub

'@EntryPoint
'@Description("Starts the loop through all subfolders of base folder.")
Public Sub LoopAllSubFoldersStart()
Attribute LoopAllSubFoldersStart.VB_Description = "Starts the loop through all subfolders of base folder."
    Dim FSOLibrary As Object
    Set FSOLibrary = CreateObject("Scripting.FileSystemObject")
    ' Start recursive loop in base folder.
    LoopAllSubFolders FSOLibrary.GetFolder(m_basePath)
End Sub
