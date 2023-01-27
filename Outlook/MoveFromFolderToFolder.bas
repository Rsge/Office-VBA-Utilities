Attribute VB_Name = "MoveFromFolderToFolder"
Attribute VB_Description = "Save movement of mails from one folder to another."
'@Folder("Mail.Moving")
'@ModuleDescription("Save movement of mails from one folder to another.")
Option Explicit

' Sleep
'@Description("Pauses the program for specified amount of milliseconds.")
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
Attribute Sleep.VB_Description = "Pauses the program for specified amount of milliseconds."

' String constants
'@VariableDescription("Name of mailbow to run this script on.")
Private Const m_mailboxName As String = "test@example.com"
Attribute m_mailboxName.VB_VarDescription = "Name of mailbow to run this script on."
'@VariableDescription("Name of superfolder to the folder to move from.")
Private Const m_superFolderName As String = "Inbox"
Attribute m_superFolderName.VB_VarDescription = "Name of superfolder to the folder to move from."
'@VariableDescription("Name of folder from which to move emails.")
Private Const m_fromFolderName As String = "Example"
Attribute m_fromFolderName.VB_VarDescription = "Name of folder from which to move emails."
'@VariableDescription("Name of folder to which to move emails.")
Private Const m_toFolderName As String = "Archive"
Attribute m_toFolderName.VB_VarDescription = "Name of folder to which to move emails."
'@VariableDescription("Name of namespace to use.")
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Moves all mails from one folder to another folder reliably.")
Public Sub MoveAllMailsFromFolder()
Attribute MoveAllMailsFromFolder.VB_Description = "Moves all mails from one folder to another folder reliably."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim mailbox As Outlook.Folder
    Set mailbox = ns.Folders.Item(m_mailboxName)
    Dim inputFolder As Outlook.Folder
    Set inputFolder = mailbox.Folders.Item(m_superFolderName).Folders.Item(m_fromFolderName)
    Dim outputFolder As Outlook.Folder
    Set outputFolder = mailbox.Folders.Item(m_toFolderName)
    Dim mails As Outlook.Items
    Set mails = inputFolder.Items
    Dim i As LongLong

    For i = mails.Count To 1 Step -1
        mails.Item(i).Move outputFolder
        Sleep 200
        DoEvents
    Next
End Sub
