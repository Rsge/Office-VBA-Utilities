Attribute VB_Name = "BulkEditContacts"
Attribute VB_Description = "Contact bulk-editing."
'@Folder "Contact editing"
'@ModuleDescription "Contact bulk-editing."
Option Explicit

'String constant
'@VariableDescription "Name of namespace to use."
Private Const NSName As String = "MAPI"
Attribute NSName.VB_VarDescription = "Name of namespace to use."


'@Description "Validates if a string starts with a given string."
Private Function StartsWith(ByVal Text As String, ByVal Beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
    StartsWith = (Left$(Trim$(UCase$(Text)), Len(Beginning)) = UCase$(Beginning))
End Function

'@EntryPoint
'@Description "Edits all contacts with changes specified in code."
Public Sub EditContacts()
Attribute EditContacts.VB_Description = "Edits all contacts with changes specified in code."
    Dim NS As Outlook.Namespace
    Set NS = GetNamespace(NSName)
    Dim ContactsFolder As Outlook.MAPIFolder
    Set ContactsFolder = NS.GetDefaultFolder(olFolderContacts)
    Dim Contacts As Outlook.Items
    Set Contacts = ContactsFolder.Items
    Dim Contact As Outlook.ContactItem
    On Error Resume Next
    For Each Contact In Contacts
        'Make changes to contact
        With Contact
            Const SearchFor As String = "test@"
            Const ReplaceWith As String = "mailto:test@"
            If StartsWith(.WebPage, SearchFor) Then
                .WebPage = Replace(.WebPage, SearchFor, ReplaceWith)
                .Close (olSave)
            End If
        End With
    Next
    On Error GoTo 0
End Sub
