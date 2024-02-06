VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteSentFromSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Deletion of emails of specific sender from Sent folder."
'@Folder("Mail.Moving")
'@ModuleDescription("Deletion of emails of specific sender from Sent folder.")
Option Explicit

' String constants
'@VariableDescription("Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder.")
Private Const m_senderAddressStart As String = "test@"
Attribute m_senderAddressStart.VB_VarDescription = "Start of sender's email address for whom to move sent mails to given mailbox's 'Sent items' folder."
'@VariableDescription("Name of namespace to use.")
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

' ————————————————————————————————————————————————————— '


'@Description("Validates if a string starts with a given string.")
Public Function StartsWith(ByVal Text As String, ByVal beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
     StartsWith = (Left$(Trim$(UCase$(Text)), Len(beginning)) = UCase$(beginning))
End Function

'@Description("On application startup, gets items in Sent and Deleted folder and deletes all mails from the specified sender.")
Private Sub Application_Startup()
Attribute Application_Startup.VB_Description = "On application startup, gets items in Sent and Deleted folder and deletes all mails from the specified sender."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(m_nsName)
    Dim folders(1) As Outlook.MAPIFolder
    Set folders(0) = NS.GetDefaultFolder(olFolderSentMail)
    Set folders(1) = NS.GetDefaultFolder(olFolderDeletedItems)
    Dim mails As Outlook.Items
    Dim i As Long
    Dim j As LongLong
    For i = 0 To UBound(folders)
        Set mails = folders(i).Items
        For j = mails.Count To 1 Step -1
            On Error Resume Next
            If TypeOf mails.Item(j) Is Outlook.MailItem Then
                Dim from As String
                from = mails.Item(j).sender.GetExchangeUser().PrimarySmtpAddress
                If StartsWith(from, m_senderAddressStart) Then
                    mails.Item(j).Delete
                End If
            End If
            On Error GoTo 0
            Sleep 100
            DoEvents
        Next
    Next
End Sub
