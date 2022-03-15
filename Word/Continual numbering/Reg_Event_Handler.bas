Attribute VB_Name = "Reg_Event_Handler"
Attribute VB_Description = "Event handler definition."
'@IgnoreModule SelfAssignedDeclaration
'@Folder "Continual numbering"
'@ModuleDescription "Event handler definition."
Option Explicit

'@Description "Registers the event class for handling."
Public Sub Register_Event_Handler()
Attribute Register_Event_Handler.VB_Description = "Registers the event class for handling."
    Dim WordEvent As New EventClassModule
    Set WordEvent.App = Word.Application
End Sub
