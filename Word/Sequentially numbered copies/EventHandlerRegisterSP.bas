Attribute VB_Name = "EventHandlerRegisterSP"
Attribute VB_Description = "Event handler definition."
'@Folder("SequentiallyNumberedCopies")
'@ModuleDescription("Event handler definition.")
Option Explicit
Option Private Module

'@Ignore MoveFieldCloserToUsage
'@VariableDescription("Saves the print event for the event handler.")
Private m_wordEvent As New PrintEventSinglePage
Attribute m_wordEvent.VB_VarDescription = "Saves the print event for the event handler."

'@Description("Registers the event class for handling.")
Public Sub RegisterEventHandler()
Attribute RegisterEventHandler.VB_Description = "Registers the event class for handling."
    Set m_wordEvent.App = Word.Application
End Sub
