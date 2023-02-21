Attribute VB_Name = "Reg_Event_Handler"
Attribute VB_Description = "Event handler definition."
'@Folder("ContinualNumbering")
'@ModuleDescription("Event handler definition.")
Option Explicit
Option Private Module

'@Description("Registers the event class for handling.")
Public Sub Register_Event_Handler()
Attribute Register_Event_Handler.VB_Description = "Registers the event class for handling."
    Dim wordEvent As EventClassModule
    Set wordEvent = New EventClassModule
    Set wordEvent.App = Word.Application
End Sub
