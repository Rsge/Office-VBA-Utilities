VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveToCorrectSentFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Movement of emails into different Sent folder."
'@IgnoreModule DefaultMemberRequired
'@Folder "Email moving"
'@ModuleDescription "Movement of emails into different Sent folder."
Option Explicit

'String constants
'@VariableDescription "Name of mailbox in which to put sent emails from given sender."
Private Const MailboxName As String = "test@example.com"
Attribute MailboxName.VB_VarDescription = "Name of mailbox in which to put sent emails from given sender."
'@VariableDescription "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
Private Const SenderAddressStart As String = "test@"
Attribute SenderAddressStart.VB_VarDescription = "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
'@VariableDescription "Name of 'Sent items' folder."
Private Const SentItemsFolderName As String = "Sent elements"
Attribute SentItemsFolderName.VB_VarDescription = "Name of 'Sent items' folder."
'@VariableDescription "Name of namespace to use."
Private Const NSName As String = "MAPI"
Attribute NSName.VB_VarDescription = "Name of namespace to use."

'Variable
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1


'@Description "Validates if a string starts with a given string."
Public Function StartsWith(ByVal Text As String, ByVal Beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
     StartsWith = (Left$(Trim$(UCase$(Text)), Len(Beginning)) = UCase$(Beginning))
End Function

'@Description "On application startup, gets items in Sent folder."
Private Sub Application_Startup()
Attribute Application_Startup.VB_Description = "On application startup, gets items in Sent folder."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(NSName)
    Dim SentItemsFolder As Outlook.MAPIFolder
    Set SentItemsFolder = NS.GetDefaultFolder(olFolderSentMail)
    Set Items = SentItemsFolder.Items
End Sub

'@Description "On adding an item to the Sent items, moves it to the correct Sent folder."
Private Sub Items_ItemAdd(ByVal Item As Object)
Attribute Items_ItemAdd.VB_Description = "On adding an item to the Sent items, moves it to the correct Sent folder."
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then
        Dim From As String
        From = Item.Sender.GetExchangeUser().PrimarySmtpAddress
        If StartsWith(From, SenderAddressStart) Then
            Dim NS As Outlook.Namespace
            Set NS = Application.GetNamespace(NSName)
            Dim Mailbox As Outlook.Folder
            Set Mailbox = NS.Folders(MailboxName)(0)
            Dim SentItemsFolder As Outlook.Folder
            Set SentItemsFolder = Mailbox.Folders(SentItemsFolderName)(0)
            Item.Move SentItemsFolder
        End If
    End If
    On Error GoTo 0
End Sub


''@Description "On sending an item, saves a copy to a different folder and deletes the original."
'Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'    'String constants
'    Const Prompt As String = "Are you sure you want to send "
'    Const QuestionMark As String = "?"
'    Const Hint As String = "Hint"
'
'    On Error Resume Next
'    Dim Namespace As Outlook.Namespace
''    Dim Mailbox As Outlook.Folder
'    Dim Mailbox As Outlook.Recipient
'    Dim SentItemsFolder As Outlook.Folder
'    Dim Mail As Outlook.MailItem
'    Dim From As String
'
'    Set Namespace = Application.GetNamespace(NSName)
'    Set Mailbox = Namespace.Folders(MailboxName)
''    Set Mailbox = Namespace.CreateRecipient(MailboxName)
'    Set SentItemsFolder = Mailbox.Folders(SentItemsFolderName)
''    Set SentItemsFolder = Namespace.GetSharedDefaultFolder(Mailbox, olFolderSentMail)
'    From = Item.SentOnBehalfOfName
''    Debug.Print (Item.SendUsingAccount)
''    Debug.Print (From)
''    Debug.Print (SentItemsFolder)
'
'    If StartsWith(From, ReceiverAddressStart) Then
''        Set Item.SaveSentMessageFolder = SentItemsFolder
'        Dim Copy As Outlook.MailItem
'        Set Copy = Item.Copy
'        Copy.Move (SentItemsFolder)
'        Item.DeleteAfterSubmit = True
'    End If
'
'    Dim Prompt As String
'    Prompt = Prompt & Item.Subject & QuestionMark
'    If MsgBox(Prompt, vbYesNo + vbQuestion, Hint) = vbNo Then
'        Cancel = True
'    End If
'    On Error GoTo 0
'End Sub
