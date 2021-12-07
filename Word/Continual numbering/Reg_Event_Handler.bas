Attribute VB_Name = "Reg_Event_Handler"
Option Explicit

Dim X As New EventClassModule
Sub Register_Event_Handler()
    Set X.App = Word.Application
End Sub
