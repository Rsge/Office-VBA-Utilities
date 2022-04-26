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
Private Const m_mailboxName As String = "test@example.com"
Attribute m_mailboxName.VB_VarDescription = "Name of mailbox in which to put sent emails from given sender."
'@VariableDescription "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
Private Const m_senderAddressStart As String = "test@"
Attribute m_senderAddressStart.VB_VarDescription = "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
'@VariableDescription "Name of 'Sent items' folder."
Private Const m_sentItemsFolderName As String = "Sent elements"
Attribute m_sentItemsFolderName.VB_VarDescription = "Name of 'Sent items' folder."
'@VariableDescription "Name of namespace to use."
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

'Variable
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents m_items As Outlook.Items
Attribute m_items.VB_VarHelpID = -1


'@Description "Validates if a string starts with a given string."
Public Function StartsWith(ByVal text As String, ByVal beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
     StartsWith = (Left$(Trim$(UCase$(text)), Len(beginning)) = UCase$(beginning))
End Function

'@Description "On application startup, gets items in Sent folder."
Private Sub Application_Startup()
Attribute Application_Startup.VB_Description = "On application startup, gets items in Sent folder."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim sentItemsFolder As Outlook.MAPIFolder
    Set sentItemsFolder = ns.GetDefaultFolder(olFolderSentMail)
    Set m_items = sentItemsFolder.Items
End Sub

'@Description "On adding an item to the Sent items, moves it to the correct Sent folder."
Private Sub m_items_ItemAdd(ByVal itemObject As Object)
Attribute m_items_ItemAdd.VB_Description = "On adding an item to the Sent items, moves it to the correct Sent folder."
    On Error Resume Next
    If TypeOf itemObject Is Outlook.MailItem Then
        Dim from As String
        from = itemObject.sender.GetExchangeUser().PrimarySmtpAddress
        If StartsWith(from, m_senderAddressStart) Then
            Dim ns As Outlook.Namespace
            Set ns = Application.GetNamespace(m_nsName)
            Dim mailbox As Outlook.Folder
            Set mailbox = ns.Folders(m_mailboxName)(0)
            Dim sentItemsFolder As Outlook.Folder
            Set sentItemsFolder = mailbox.Folders(m_sentItemsFolderName)(0)
            itemObject.Move sentItemsFolder
        End If
    End If
    On Error GoTo 0
End Sub


''@Description "On sending an item, saves a copy to a different folder and deletes the original."
'Private Sub Application_ItemSend(ByVal item As Object, cancel As Boolean)
'    'String constants
'    Const prompt As String = "Are you sure you want to send "
'    Const questionMark As String = "?"
'    Const hint As String = "Hint"
'
'    On Error Resume Next
'    Dim ns As Outlook.Namespace
''    Dim Mailbox As Outlook.Folder
'    Dim mailbox As Outlook.Recipient
'    Dim sentItemsFolder As Outlook.Folder
'    Dim mail As Outlook.MailItem
'    Dim from As String
'
'    Set ns = Application.GetNamespace(NSName)
'    Set mailbox = ns.Folders(MailboxName)
''    Set Mailbox = Namespace.CreateRecipient(MailboxName)
'    Set sentItemsFolder = mailbox.Folders(SentItemsFolderName)
''    Set SentItemsFolder = Namespace.GetSharedDefaultFolder(Mailbox, olFolderSentMail)
'    from = item.SentOnBehalfOfName
''    Debug.Print (Item.SendUsingAccount)
''    Debug.Print (From)
''    Debug.Print (SentItemsFolder)
'
'    If StartsWith(from, m_senderAddressStart) Then
''        Set item.SaveSentMessageFolder = sentItemsFolder
'        Dim copy As Outlook.MailItem
'        Set copy = item.copy
'        copy.Move (sentItemsFolder)
'        item.DeleteAfterSubmit = True
'    End If
'
'    Dim prompt As String
'    prompt = prompt & item.subject & questionMark
'    If MsgBox(prompt, vbYesNo + vbQuestion, hint) = vbNo Then
'        cancel = True
'    End If
'    On Error GoTo 0
'End Sub
