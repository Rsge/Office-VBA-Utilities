Attribute VB_Name = "Unification"
Attribute VB_Description = "Handles unification of certain aspects in all files of a type in a starting folder and all it's subfolders."
'@Folder("FileEditing")
'@IgnoreModule ImplicitDefaultMemberAccess
'@ModuleDescription("Handles unification of certain aspects in all files of a type in a starting folder and all it's subfolders.")
Option Explicit

'@VariableDescription("File extension of file type to process.")
Private Const m_fileExt As String = ".docx"
Attribute m_fileExt.VB_VarDescription = "File extension of file type to process."
'@VariableDescription("Starting folder of document search.")
Private Const m_basePath As String = "C:\Example"
Attribute m_basePath.VB_VarDescription = "Starting folder of document search."
'@VariableDescription("Name of the 'Standard' format template.")
Private Const m_standardStyleName As String = "Standard"
Attribute m_standardStyleName.VB_VarDescription = "Name of the 'Standard' format template."
'@VariableDescription("Font of the 'Standard' format template.")
Private Const m_standardFont As String = "Arial"
Attribute m_standardFont.VB_VarDescription = "Font of the 'Standard' format template."
'@VariableDescription("Font size of the 'Standard' format template.")
Private Const m_standardFontSize As Long = 12
Attribute m_standardFontSize.VB_VarDescription = "Font size of the 'Standard' format template."
'@VariableDescription("Name of the 'Heading 1' format template.")
Private Const m_heading1StyleName As String = "Heading 1"
Attribute m_heading1StyleName.VB_VarDescription = "Name of the 'Heading 1' format template."
'@VariableDescription("Font of the 'Heading 1' format template.")
Private Const m_heading1Font As String = "Arial"
Attribute m_heading1Font.VB_VarDescription = "Font of the 'Heading 1' format template."
'@VariableDescription("Font size of the 'Heading 1' format template.")
Private Const m_heading1FontSize As Long = 14
Attribute m_heading1FontSize.VB_VarDescription = "Font size of the 'Heading 1' format template."
'@VariableDescription("If the font of the 'Heading 1' format template is bold.")
Private Const m_heading1Bold As Boolean = False
Attribute m_heading1Bold.VB_VarDescription = "If the font of the 'Heading 1' format template is bold."
'@VariableDescription("If the font of the 'Heading 1' format template is italic.")
Private Const m_heading1Italic As Boolean = False
Attribute m_heading1Italic.VB_VarDescription = "If the font of the 'Heading 1' format template is italic."
'@VariableDescription("How the font of the 'Heading 1' format template is underlined.")
Private Const m_heading1Unterline As Long = wdUnderlineSingle
Attribute m_heading1Unterline.VB_VarDescription = "How the font of the 'Heading 1' format template is underlined."

' ————————————————————————————————————————————————————— '


'@Description("Tests if a string ends with another string.")
Private Function EndsWith(ByVal str As String, ByVal ending As String) As Boolean
Attribute EndsWith.VB_Description = "Tests if a string ends with another string."
    EndsWith = Right$(str, Len(ending)) = ending
End Function

'@Description("Makes changes in a file via regex replace.")
Private Sub UnifyStyles(ByVal filePath As String, Optional ByVal closeAfter As Boolean = True)
Attribute UnifyStyles.VB_Description = "Makes changes in a file via regex replace."
    ' Open doc.
    Dim doc As Document
    Set doc = Documents.Open(filePath)
    Dim changed As Boolean
    changed = False
    ' Unify Style.
    With ActiveDocument.Styles.Item(m_standardStyleName).Font
        If .Name <> m_standardFont Then
            .Name = m_standardFont
            .Size = m_standardFontSize
            changed = True
        End If
    End With
    With ActiveDocument.Styles.Item(m_heading1StyleName).Font
        If .Name <> m_heading1Font Then
            .Name = m_heading1Font
            .Size = m_heading1FontSize
            .Bold = m_heading1Bold
            .Italic = m_heading1Italic
            .Underline = m_heading1Unterline
            .Color = wdColorAutomatic
            changed = True
        End If
    End With
    ' Save and close file.
    If changed Then
        doc.Save
    End If
    If closeAfter Then
        doc.Close wdDoNotSaveChanges
    End If
End Sub

'@Description("Makes changes in a file via regex replace.")
Private Sub UnifyHeaders(ByVal filePath As String, ByRef i As Long, ByRef formattedTxt As Range, ByRef tabs As TabStops, _
                         Optional ByVal closeAfter As Boolean = True)
Attribute UnifyHeaders.VB_Description = "Makes changes in a file via regex replace."
    ' Open doc.
    Dim doc As Document
    Set doc = Documents.Open(filePath)
    Dim changed As Boolean
    changed = False
    ' Replace header in file.
    Dim headerRange As Range
    Set headerRange = doc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range
    If Len(headerRange.Text) > 4 Then ' If there even is a header there...
        If i = 1 Then
            Set formattedTxt = headerRange.FormattedText
            Set tabs = headerRange.ParagraphFormat.TabStops
        ElseIf headerRange.FormattedText <> formattedTxt Then
            headerRange.FormattedText = formattedTxt
            '@Ignore ValueRequired
            headerRange.ParagraphFormat.TabStops = tabs
            headerRange.FormattedText.Characters.Last.Previous.Text = vbNullString ' Delete trailing newline from insert.
            changed = True
        End If
        ' Save and close file.
        If changed Then
            doc.Save
        End If
        DoEvents
        If i > 1 And closeAfter Then
            doc.Close wdDoNotSaveChanges
        End If
        i = i + 1
    ElseIf closeAfter Then
        doc.Close wdDoNotSaveChanges
    End If
End Sub

'@Description("Loops through all subfolders of base folder and through all their files.")
Private Sub LoopAllSubFolders(ByVal FSOFolder As Object, ByRef i As Long, ByRef formattedTxt As Range, ByRef tabs As TabStops)
Attribute LoopAllSubFolders.VB_Description = "Loops through all subfolders of base folder and through all their files."
    Dim FSOSubFolder As Object
    Dim FSOFile As Object
    ' Recurse into subfolders.
    For Each FSOSubFolder In FSOFolder.Subfolders
        LoopAllSubFolders FSOSubFolder, i, formattedTxt, tabs
    Next
    ' For each file of a specified type in a subfolder, change stuff.
    For Each FSOFile In FSOFolder.Files
        If EndsWith(FSOFile.Name, m_fileExt) And InStrB(FSOFile.Name, "~$") <> 1 Then
            UnifyStyles FSOFile.Path, closeAfter:=False
            UnifyHeaders FSOFile.Path, i, formattedTxt, tabs
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
    Dim i As Long
    i = 1
    Dim formattedTxt As Range
    Set formattedTxt = Nothing
    Dim tabs As TabStops
    Set tabs = Nothing
    ' Start recursive loop in base folder.
    LoopAllSubFolders FSOLibrary.GetFolder(m_basePath), i, formattedTxt, tabs
End Sub
