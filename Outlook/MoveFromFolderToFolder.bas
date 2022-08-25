Attribute VB_Name = "MoveFromFolderToFolder"
Attribute VB_Description = "Save movement of mails from one folder to another."
'@IgnoreModule DefaultMemberRequired
'@Folder "Email moving"
'@ModuleDescription "Save movement of mails from one folder to another."
Option Explicit

'String constants
'@VariableDescription "Name of mailbow to run this script on."
Private Const m_mailboxName As String = "test@example.com"
Attribute m_mailboxName.VB_VarDescription = "Name of mailbow to run this script on."
'@VariableDescription "Name of superfolder to the folder to move from."
Private Const m_superFolderName As String = "Inbox"
Attribute m_superFolderName.VB_VarDescription = "Name of superfolder to the folder to move from."
'@VariableDescription "Name of folder from which to move emails."
Private Const m_fromFolderName As String = "Example"
Attribute m_fromFolderName.VB_VarDescription = "Name of folder from which to move emails."
'@VariableDescription "Name of folder to which to move emails."
Private Const m_toFolderName As String = "Archive"
Attribute m_toFolderName.VB_VarDescription = "Name of folder to which to move emails."
'@VariableDescription "Name of namespace to use."
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."


'@EntryPoint
'@Description "Moves all mails from one folder to another folder reliably."
Public Sub MoveAllMailsFromFolder()
Attribute MoveAllMailsFromFolder.VB_Description = "Moves all mails from one folder to another folder reliably."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim mailbox As Outlook.Folder
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set mailbox = ns.Folders(m_mailboxName)
    Dim inputFolder As Outlook.Folder
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set inputFolder = mailbox.Folders(m_superFolderName).Folders(m_fromFolderName)
    Dim outputFolder As Outlook.Folder
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set outputFolder = mailbox.Folders(m_toFolderName)
    Dim mails As Outlook.Items
    Set mails = inputFolder.Items
    Dim i As LongLong

    For i = mails.Count To 1 Step -1
        mails.Item(i).Move outputFolder
        DoEvents
    Next
End Sub
