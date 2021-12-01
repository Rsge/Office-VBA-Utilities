VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveToCorrectSentFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1

Public Function StartsWith(Text As String, Beginning As String) As Boolean
     StartsWith = (Left(Trim(UCase(Text)), Len(Beginning)) = UCase(Beginning))
End Function

Private Sub Application_Startup()
    Dim NS As Outlook.Namespace
    Dim SentItemsFolder  As Outlook.MAPIFolder

    Set NS = Application.GetNamespace("MAPI")
    Set SentItemsFolder = NS.GetDefaultFolder(olFolderSentMail)
    Set Items = SentItemsFolder.Items
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        Dim From As String
        From = Item.Sender.GetExchangeUser().PrimarySmtpAddress
        Debug.Print (From)
        If StartsWith(From, "test@") Then
            On Error Resume Next
            Dim Namespace As Outlook.Namespace
            Dim Mailbox As Outlook.Folder
            Dim SentItemsFolder As Outlook.Folder
            Set Namespace = Application.GetNamespace("MAPI")
            Set Mailbox = Namespace.Folders("test@example.com")
            Set SentItemsFolder = Mailbox.Folders("Sent elements") 'Name of your sent items folder

            Item.Move SentItemsFolder
        End If
    End If
End Sub


'Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
''    On Error Resume Next
'    Dim Namespace As Outlook.Namespace
''    Dim Mailbox As Outlook.Folder
'    Dim Mailbox As Outlook.Recipient
'    Dim SentItemsFolder As Outlook.Folder
'    Dim Mail As Outlook.MailItem
'    Dim From As String
'
'    Set Namespace = Application.GetNamespace("MAPI")
'    Set Mailbox = Namespace.Folders("test@example.com")
''    Set Mailbox = Namespace.CreateRecipient("test@example.com")
'    Set SentItemsFolder = Mailbox.Folders("Sent elements")
''    Set SentItemsFolder = Namespace.GetSharedDefaultFolder(Mailbox, olFolderSentMail)
'    From = Item.SentOnBehalfOfName
''    Debug.Print (Item.SendUsingAccount)
''    Debug.Print (From)
''    Debug.Print (SentItemsFolder)
'
'    If StartsWith(From, "test@") Then
''        Set Item.SaveSentMessageFolder = SentItemsFolder
'        Dim Kopie As Outlook.MailItem
'        Set Kopie = Item.Copy
'        Kopie.Move (SentItemsFolder)
'        Item.DeleteAfterSubmit = True
'    End If
'
'    Dim prompt As String
'    prompt = "Are you sure you want to send " & Item.Subject & "?"
'    If MsgBox(prompt, vbYesNo + vbQuestion, "Sample") = vbNo Then
'        Cancel = True
'    End If
'    Err.Clear
'End Sub




