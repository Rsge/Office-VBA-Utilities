Attribute VB_Name = "MoveFromFolderToFolder"
Attribute VB_Description = "Module for methods related to moving all mails from one folder to another savely."
'@IgnoreModule DefaultMemberRequired
'@Folder "Email moving"
'@ModuleDescription "Module for methods related to moving all mails from one folder to another savely."
Option Explicit

'String constants
Private Const MailboxName As String = "test@example.com"
Private Const FromFolderName As String = "Archive"
Private Const ToFolderName As String = "Test"
Private Const MAPI As String = "MAPI"


'@EntryPoint
'@Description "Move all mails from one folder to another folder reliably."
Public Sub MoveAllMailsFromFolder()
Attribute MoveAllMailsFromFolder.VB_Description = "Move all mails from one folder to another folder reliably."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(MAPI)
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
