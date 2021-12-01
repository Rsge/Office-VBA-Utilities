Attribute VB_Name = "JunkMailBlackWhitelist"
Option Explicit
Sub WhitelistMailsFromWebsite()
    Dim Namespace As Outlook.Namespace
    Dim Inbox As Outlook.Folder
    Dim Junk As Outlook.Folder
    Dim Mail As Outlook.MailItem
    Dim Subject As String
    
    Set Namespace = Application.GetNamespace("MAPI")
    Set Inbox = Namespace.GetDefaultFolder(olFolderInbox)
    Set Junk = Namespace.GetDefaultFolder(olFolderJunk)
    
    For Each Mail In Junk.Items
        Subject = Mail.Subject
        
        If InStr(Subject, "Test") > 0 Then
            Mail.Move (Inbox)
        End If
    Next
End Sub

Sub JunkMailFilter()
    Dim Namespace As Outlook.Namespace
    Dim Inbox As Outlook.Folder
    Dim Junk As Outlook.Folder
    Dim Mail As Outlook.MailItem
    Dim MailObjekt As Object
    Dim StringObjekt As Variant
    
    Dim FileName As String
    Dim FileNumber As Integer
    Dim MailList As Object
    Dim MailAddress As String
    
    Set Namespace = Application.GetNamespace("MAPI")
    Set Inbox = Namespace.GetDefaultFolder(olFolderInbox)
    Set Junk = Namespace.GetDefaultFolder(olFolderJunk)
    Set MailList = CreateObject("System.Collections.ArrayList")
    
    FileNumber = FreeFile()
    FileName = "Path\To\List\List.txt" 'Change this to your file's path
    Open FileName For Input As FileNumber
        
    Do While Not EOF(FileNumber)
        Line Input #FileNumber, Mailadresse
        MailList.Add Mailadresse
    Loop
    
    For Each MailObjekt In Inbox.Items
        If (TypeOf MailObjekt Is Outlook.MailItem) Then
            Set Mail = MailObjekt
            For Each StringObjekt In MailList
                If TypeName(StringObjekt) = "String" Then
                    Mailadresse = StringObjekt
                    If FindMailAddress(Mail) = Mailadresse Then
                        Mail.Move (Junk)
                        Exit For
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function FindMailAddress(Mail As Outlook.MailItem) As String
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


