Attribute VB_Name = "MoveFromFolderToFolder"
Option Explicit

Sub MoveAllMailsFromFolder()
    Dim Namespace As Outlook.Namespace
    Dim Mailbox As Outlook.Folder
    Dim InputFolder As Outlook.Folder
    Dim OutputFolder As Outlook.Folder
    Dim Mails As Outlook.Items
    Dim Mail As Object
    Dim i As Long
    
    Set Namespace = Application.GetNamespace("MAPI")
    Set Mailbox = Namespace.Folders("test@example.com")
    Set InputFolder = Mailbox.Folders("Archive")
    Set OutputFolder = Mailbox.Folders("Test")
    Set Mails = InputFolder.Items
    
    'Debug.Print (Mailbox)
    'Debug.Print (InputFolder)
    'Debug.Print (OutputFolder)

    For i = Mails.Count To 1 Step -1
        Mails.Item(i).Move OutputFolder
    Next
End Sub
