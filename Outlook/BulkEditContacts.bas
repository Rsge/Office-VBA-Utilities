Attribute VB_Name = "BulkEditContacts"
Attribute VB_Description = "Contact bulk-editing."
'@Folder "Contact editing"
'@ModuleDescription "Contact bulk-editing."
Option Explicit

'String constant
'@VariableDescription "Term to search for in contact."
Private Const m_searchFor As String = "test@"
Attribute m_searchFor.VB_VarDescription = "Term to search for in contact."
'@VariableDescription "Term to replace searched term with."
Private Const m_replaceWith As String = "mailto:test@"
Attribute m_replaceWith.VB_VarDescription = "Term to replace searched term with."
'@VariableDescription "Name of namespace to use."
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."


'@Description "Validates if a string starts with a given string."
Private Function StartsWith(ByVal text As String, ByVal beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
    StartsWith = (Left$(Trim$(UCase$(text)), Len(beginning)) = UCase$(beginning))
End Function

'@EntryPoint
'@Description "Edits all contacts with changes specified in code."
Public Sub EditContacts()
Attribute EditContacts.VB_Description = "Edits all contacts with changes specified in code."
    Dim ns As Outlook.Namespace
    Set ns = GetNamespace(m_nsName)
    Dim contactsFolder As Outlook.MAPIFolder
    Set contactsFolder = ns.GetDefaultFolder(olFolderContacts)
    Dim contacts As Outlook.Items
    Set contacts = contactsFolder.Items
    Dim contact As Outlook.ContactItem
    On Error Resume Next
    For Each contact In contacts
        'Making changes to contact
        With contact
            If StartsWith(.WebPage, m_searchFor) Then
                .WebPage = Replace(.WebPage, m_searchFor, m_replaceWith)
                .Close (olSave)
            End If
        End With
    Next
    On Error GoTo 0
End Sub
