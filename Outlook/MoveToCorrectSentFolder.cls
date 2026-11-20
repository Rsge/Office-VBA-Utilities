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
'@Folder("Mail.Moving")
'@ModuleDescription("Movement of emails into different Sent folder.")
Option Explicit

' String constants
'@VariableDescription("Name of mailbox in which to put sent emails from given sender.")
Private Const m_extMailboxName As String = "Example Test"
Attribute m_extMailboxName.VB_VarDescription = "Name of mailbox in which to put sent emails from given sender."
'@VariableDescription("Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder.")
Private Const m_senderAddressStart As String = "test@"
Attribute m_senderAddressStart.VB_VarDescription = "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
'@VariableDescription("Name of 'Sent items' folder.")
Private Const m_sentFolderName As String = "Sent elements"
Attribute m_sentFolderName.VB_VarDescription = "Name of 'Sent items' folder."
'@VariableDescription("Name of namespace to use.")
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

' Variables
'@MemberAttribute VB_VarHelpID, -1
'@VariableDescription("Items in sent folder with events.")
Private WithEvents m_sentItems As Outlook.Items
Attribute m_sentItems.VB_VarDescription = "Items in sent folder with events."
'@VariableDescription("External sent items folder.")
Private m_extSentFolder As Outlook.Folder
Attribute m_extSentFolder.VB_VarDescription = "External sent items folder."

' ————————————————————————————————————————————————————— '


'@Description("Validates if a string starts with a given string.")
Private Function StartsWith(ByVal Text As String, ByVal beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
     StartsWith = (Left$(Trim$(UCase$(Text)), Len(beginning)) = UCase$(beginning))
End Function

'@Description("On application startup, gets items in Sent folder.")
Private Sub Application_Startup()
Attribute Application_Startup.VB_Description = "On application startup, gets items in Sent folder."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim extMailbox As Outlook.Folder
    Set extMailbox = ns.Folders.Item(m_extMailboxName)
    Set m_extSentFolder = extMailbox.Folders.Item(m_sentFolderName)
    Dim defaultSentFolder As Outlook.Folder
    Set defaultSentFolder = ns.GetDefaultFolder(olFolderSentMail)
    Set m_sentItems = m_defaultSentFolder.Items
End Sub

'@Description("On adding an item to the Sent items, moves it to the correct Sent folder.")
Private Sub m_sentItems_ItemAdd(ByVal Item As Object)
Attribute m_sentItems_ItemAdd.VB_Description = "On adding an item to the Sent items, moves it to the correct Sent folder."
    If TypeOf Item Is Outlook.MailItem Then
        Dim from As String
        from = Item.sender.GetExchangeUser().PrimarySmtpAddress
        If StartsWith(from, m_senderAddressStart) Then
            On Error Resume Next
            Item.Move m_extSentFolder
            On Error Goto 0
        End If
    End If
End Sub