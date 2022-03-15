Attribute VB_Name = "JunkMailBlackWhitelist"
'@Folder "Junk mail handling"
'@ModuleDescription "Module for methods related to more extensive whitelisting and blacklisting of emails.
Option Explicit

'String constants
Private Const WhitelistSubject As String = "Test"
Private Const JunkMailListPath As String = "C:\Path\To\List\List.txt"
Private Const MAPI As String = "MAPI"


'@EntryPoint
'@Description "Remove all mails with a given subject from Junk."
Public Sub WhitelistMailsWithSubject()
Attribute WhitelistMailsWithSubject.VB_Description = "Remove all mails with a given subject from Junk."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(MAPI)
    Dim Inbox As Outlook.Folder
    Set Inbox = NS.GetDefaultFolder(olFolderInbox)
    Dim Junk As Outlook.Folder
    Set Junk = NS.GetDefaultFolder(olFolderJunk)
    Dim Mail As Outlook.MailItem
    Dim Subject As String
    
    For Each Mail In Junk.Items
        Subject = Mail.Subject
        If InStr(Subject, WhitelistSubject) > 0 Then
            Mail.Move Inbox
        End If
    Next
End Sub

'@EntryPoint
'@Description "Move all Emails on a specified TXT-List into Junk."
Public Sub JunkMailFilter()
Attribute JunkMailFilter.VB_Description = "Move all Emails on a specified TXT-List into Junk."
    Dim NS As Outlook.Namespace
    Set NS = Application.GetNamespace(MAPI)
    Dim Inbox As Outlook.Folder
    Set Inbox = NS.GetDefaultFolder(olFolderInbox)
    Dim Junk As Outlook.Folder
    Set Junk = NS.GetDefaultFolder(olFolderJunk)
    Dim Mail As Outlook.MailItem
    Dim MailObject As Object
    Dim StringObject As Variant
    Dim FileName As String
    FileName = JunkMailListPath
    Dim FileNumber As Long
    FileNumber = FreeFile()
    Dim MailList As Object
    Set MailList = CreateObject("System.Collections.ArrayList")
    Dim MailAddress As String
    
    'Read File
    Open FileName For Input As FileNumber
    Do While Not EOF(FileNumber)
        Line Input #FileNumber, MailAddress
        MailList.Add MailAddress
    Loop
    Close FileName
    
    'Determine if Junk and process accordingly
    For Each MailObject In Inbox.Items
        If (TypeOf MailObject Is Outlook.MailItem) Then
            Set Mail = MailObject
            For Each StringObject In MailList
                If TypeName(StringObject) = "String" Then
                    MailAddress = StringObject
                    If FindMailAddress(Mail) = MailAddress Then
                        Mail.Move Junk
                        Exit For
                    End If
                End If
            Next
        End If
    Next
End Sub

'@Description "Find the mail address of a given mail item."
Private Function FindMailAddress(ByVal Mail As Outlook.MailItem) As String
Attribute FindMailAddress.VB_Description = "Find the mail address of a given mail item."
    If Mail.SenderEmailType = "EX" Then
        Dim Sender As Outlook.AddressEntry
        Set Sender = Mail.Sender
        If Not Sender Is Nothing Then
            If Sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or Sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                Dim ExchangeUser As Outlook.ExchangeUser
                Set ExchangeUser = Sender.GetExchangeUser()
                If Not ExchangeUser Is Nothing Then
                     FindMailAddress = ExchangeUser.PrimarySmtpAddress
                Else
                    FindMailAddress = vbNullString
                End If
            Else
                 FindMailAddress = Sender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            End If
        Else
            FindMailAddress = vbNullString
        End If
    Else
        FindMailAddress = Mail.SenderEmailAddress
    End If
End Function
