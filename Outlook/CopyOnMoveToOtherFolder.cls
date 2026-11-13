VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyOnMoveToOtherFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Copy emails moved into a folder into another folder."
'@Folder("Mail.Moving")
'@ModuleDescription("Copy emails moved into a folder into another folder.")
Option Explicit

' String constants
'@VariableDescription("Name of mailbox in which to put sent emails from given sender.")
Private Const m_extMailboxName As String = "test@example.com"
Attribute m_extMailboxName.VB_VarDescription = "Name of mailbox in which to put sent emails from given sender."
'@VariableDescription("Name of watched folder.")
Private Const m_watchedFolderName As String = "Archive"
Attribute m_watchedFolderName.VB_VarDescription = "Name of watched folder."
'@VariableDescription("Name of folder for copies.")
Private Const m_copyToFolderName As String = "CopyArchive"
Attribute m_copyToFolderName.VB_VarDescription = "Name of folder for copies."
'@VariableDescription("Name of namespace to use.")
Private Const m_nsName As String = "MAPI"
Attribute m_nsName.VB_VarDescription = "Name of namespace to use."

' Variable
'@VariableDescription("If next event shall be canceled.")
Private m_cancelEvent As Boolean
Attribute m_cancelEvent.VB_VarDescription = "If next event shall be canceled."
'@VariableDescription("Watched folder.")
Private m_watchedFolder As Outlook.Folder
Attribute m_watchedFolder.VB_VarDescription = "Watched folder."
'@VariableDescription("Folder for copies.")
Private m_copyToFolder As Outlook.Folder
Attribute m_copyToFolder.VB_VarDescription = "Folder for copies."
'@VariableDescription("Items in watched folder with events.")
Private WithEvents m_watchedItems As Outlook.Items
Attribute m_watchedItems.VB_VarDescription = "Items in sent folder with events."

' ————————————————————————————————————————————————————— '


'@Description("Validates if a string starts with a given string.")
Private Function StartsWith(ByVal Text As String, ByVal beginning As String) As Boolean
Attribute StartsWith.VB_Description = "Validates if a string starts with a given string."
     StartsWith = (Left$(Trim$(UCase$(Text)), Len(beginning)) = UCase$(beginning))
End Function

'@Description("On application startup, gets items in watched folder.")
Private Sub Application_Startup()
Attribute Application_Startup.VB_Description = "On application startup, gets items in watched folder."
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace("MAPI")
    Dim extMailbox As Outlook.Folder
    Set extMailbox = ns.Folders.Item(m_extMailboxName)
    Set m_watchedFolder = extMailbox.Folders.Item(m_watchedFolderName)
    Set m_watchedItems = m_watchedFolder.Items
    Set m_copyToFolder = extMailbox.Folders.Item(m_copyToFolderName)
    m_cancelEvent = False
End Sub

'@Description("On adding an item to the watched items, copies it to another folder.")
Private Sub m_watchedItems_ItemAdd(ByVal Item As Object)
Attribute m_watchedItems_ItemAdd.VB_Description = "On adding an item to the watched items, copies it to another folder."
    If m_cancelEvent Then
        ' Blocks infinite repeats because of the copy triggering this event again.
        ' There does not seem to be a more elegant way to do this...
        m_cancelEvent = False
        Exit Sub
    End If
    If TypeOf Item Is Outlook.MailItem Then
        On Error Resume Next
        Dim itemCopy As Outlook.MailItem
        m_cancelEvent = True
        Set itemCopy = Item.Copy
        itemCopy.Move m_copyToFolder
    End If
End Sub
