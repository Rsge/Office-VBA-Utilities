VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Open()
    Register_Event_Handler
End Sub

Public Sub DeleteVars()
    Dim Var As Variable
    For Each Var In ActiveDocument.Variables
        Var.Delete
    Next
End Sub

