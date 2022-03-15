Attribute VB_Name = "BulkEditContacts"
Attribute VB_Description = "Module for methods related to bulk-editing contacts."
'@Folder "Contact editing"
'@ModuleDescription "Module for methods related to bulk-editing contacts."
Option Explicit

'String constant
Private Const MAPI As String = "MAPI"


'@Description "Validates if a string starts with a given string."
Private Function StartsWith(ByVal Text As String, ByVal Beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
    StartsWith = (Left$(Trim$(UCase$(Text)), Len(Beginning)) = UCase$(Beginning))
End Function

'@EntryPoint
'@Description "Edit all contacts with changes specified in code."
Public Sub EditContacts()
Attribute EditContacts.VB_Description = "Edit all contacts with changes specified in code."
    On Error GoTo ErrorHandler
    Dim NS As Outlook.Namespace
    Set NS = GetNamespace(MAPI)
    Dim ContactsFolder As Outlook.MAPIFolder
    Set ContactsFolder = NS.GetDefaultFolder(olFolderContacts)
    Dim Contacts As Outlook.Items
    Set Contacts = ContactsFolder.Items
    Dim Contact As Outlook.ContactItem
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
    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub
