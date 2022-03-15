VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Event-available class."
'@Exposed
'@PredeclaredId
'@Folder "Production planning"
'@ModuleDescription "Event-available class."
Option Explicit

'@Description "On each worksheet change, update the slowdowns."
Private Sub Worksheet_Change(ByVal target As Range)
Attribute Worksheet_Change.VB_Description = "On each worksheet change, update the slowdowns."
    If Not Slowdown.IsRunning Then
        UpdateSlowdowns (target.Address)
    End If
End Sub
