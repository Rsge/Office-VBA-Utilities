Attribute VB_Name = "JunkMailBlackWhitelist"
Attribute VB_Description = "More extensive whitelisting and blacklisting of emails."
'@Folder("Mail.JunkHandling")
'@ModuleDescription("More extensive whitelisting and blacklisting of emails.")
Option Explicit

' String constants
'@VariableDescription("Subject of mail which warrants it a whitelist.")
Private Const m_whitelistSubject As String = "Test"
Attribute m_whitelistSubject.VB_VarDescription = "Subject of mail which warrants it a whitelist."
'@VariableDescription("Path to text file with list of junk-sending email addresses.")
Private Const m_junkMailListPath As String = "C:\Path\To\List\List.txt"
Attribute m_junkMailListPath.VB_VarDescription = "Path to text file with list of junk-sending email addresses."
'@VariableDescription("Name of namespace to use.")
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

' ————————————————————————————————————————————————————— '


'@Description("Finds the mail address of a given mail item.")
Private Function FindMailAddress(ByVal mail As Outlook.MailItem) As String
Attribute FindMailAddress.VB_Description = "Finds the mail address of a given mail item."
    If mail.SenderEmailType = "EX" Then
        Dim sender As Outlook.AddressEntry
        Set sender = mail.sender
        If Not sender Is Nothing Then
            If sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                Dim exchangeUser As Outlook.exchangeUser
                Set exchangeUser = sender.GetExchangeUser()
                If Not exchangeUser Is Nothing Then
                     FindMailAddress = exchangeUser.PrimarySmtpAddress
                Else
                    FindMailAddress = vbNullString
                End If
            Else
                 FindMailAddress = sender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            End If
        Else
            FindMailAddress = vbNullString
        End If
    Else
        FindMailAddress = mail.SenderEmailAddress
    End If
End Function

'@EntryPoint
'@Description("Removes all mails with a given subject from Junk.")
Public Sub WhitelistMailsWithSubject()
Attribute WhitelistMailsWithSubject.VB_Description = "Removes all mails with a given subject from Junk."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim inbox As Outlook.Folder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    Dim junk As Outlook.Folder
    Set junk = ns.GetDefaultFolder(olFolderJunk)
    Dim mail As Outlook.MailItem
    Dim subject As String
    
    For Each mail In junk.Items
        subject = mail.subject
        If InStrB(subject, m_whitelistSubject) > 0 Then
            mail.Move inbox
        End If
    Next
End Sub

'@EntryPoint
'@Description("Moves all mails on a specified TXT-List into Junk.")
Public Sub JunkMailFilter()
Attribute JunkMailFilter.VB_Description = "Moves all mails on a specified TXT-List into Junk."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace(m_nsName)
    Dim inbox As Outlook.Folder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    Dim junk As Outlook.Folder
    Set junk = ns.GetDefaultFolder(olFolderJunk)
    Dim mail As Outlook.MailItem
    Dim mailObject As Object
    Dim stringObject As Variant
    Dim fileName As String
    fileName = m_junkMailListPath
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim mailList As Object
    Set mailList = CreateObject("System.Collections.ArrayList")
    Dim mailAddress As String
    
    ' Read file.
    Open fileName For Input As fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, mailAddress
        mailList.Add mailAddress
    Loop
    Close fileName
    
    ' Determe if Junk and process accordingly.
    For Each mailObject In inbox.Items
        If (TypeOf mailObject Is Outlook.MailItem) Then
            Set mail = mailObject
            For Each stringObject In mailList
                If TypeName(stringObject) = "String" Then
                    mailAddress = stringObject
                    If FindMailAddress(mail) = mailAddress Then
                        mail.Move junk
                        Exit For
                    End If
                End If
            Next
        End If
    Next
End Sub
