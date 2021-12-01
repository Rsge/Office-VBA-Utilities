Attribute VB_Name = "BulkEditContacts"
Option Explicit

Public Function StartsWith(Text As String, Beginning As String) As Boolean
     StartsWith = (Left(Trim(UCase(Text)), Len(Beginning)) = UCase(Beginning))
End Function

Sub EditContacts()
    Dim Namespace As Outlook.Namespace
    Dim ContactsFolder As Outlook.MAPIFolder
    Dim Contact As Outlook.ContactItem
    Dim Contacts As Outlook.Items
    On Error Resume Next
    Set Namespace = GetNamespace("MAPI")
    Set ContactsFolder = Namespace.GetDefaultFolder(olFolderContacts)
    Set Contacts = ContactsFolder.Items
    For Each Contact In Contacts
        'Make changes to contact
        If StartsWith(Contact.WebPage, "test@") Then
            Debug.Print (Contact.WebPage)
            With Contact
                .WebPage = Replace(.WebPage, "test@", "mailto:test@")
                .Close (olSave)
            End With
        End If
        Err.Clear
    Next
End Sub

