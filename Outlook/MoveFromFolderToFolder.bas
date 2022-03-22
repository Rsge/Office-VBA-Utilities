Attribute VB_Name = "MoveFromFolderToFolder"
Attribute VB_Description = "Save movement of mails from one folder to another."
'@IgnoreModule DefaultMemberRequired
'@Folder "Email moving"
'@ModuleDescription "Save movement of mails from one folder to another."
Option Explicit

'String constants
'@VariableDescription "Name of mailbow to run this script on."
Private Const MailboxName As String = "test@example.com"
Attribute MailboxName.VB_VarDescription = "Name of mailbow to run this script on."
'@VariableDescription "Name of folder from which to move emails."
Private Const FromFolderName As String = "Archive"
Attribute FromFolderName.VB_VarDescription = "Name of folder from which to move emails."
'@VariableDescription "Name of folder to which to move emails."
Private Const ToFolderName As String = "Test"
Attribute ToFolderName.VB_VarDescription = "Name of folder to which to move emails."
'@VariableDescription "Name of namespace to use."
Private Const NSName As String = "MAPI"
Attribute NSName.VB_VarDescription = "Name of namespace to use."


'@EntryPoint
'@Description "Moves all mails from one folder to another folder reliably."
Public Sub MoveAllMailsFromFolder()
Attribute MoveAllMailsFromFolder.VB_Description = "Moves all mails from one folder to another folder reliably."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(NSName)
    Dim Mailbox As Outlook.Folder
    Set Mailbox = NS.Folders(MailboxName)(0)
    Dim InputFolder As Outlook.Folder
    Set InputFolder = Mailbox.Folders(FromFolderName)(0)
    Dim OutputFolder As Outlook.Folder
    Set OutputFolder = Mailbox.Folders(ToFolderName)(0)
    Dim Mails As Outlook.Items
    Set Mails = InputFolder.Items
    Dim i As Long

    For i = Mails.Count To 1 Step -1
        Mails.Item(i).Move OutputFolder
    Next
End Sub
