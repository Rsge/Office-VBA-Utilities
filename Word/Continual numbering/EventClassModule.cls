VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)
    Dim i As Long
    Dim NumCopiesToPrint As Long
    Dim Counter As Long
    Dim StartAt As Long
    
    'Ensure variables exists
    Dim VarsExist(1 To 4) As Boolean
    For i = 1 To 4
        VarsExist(i) = False
    Next
    Dim Var As Variable
    For Each Var In ActiveDocument.Variables
        For i = 1 To 4
            If Var.Name = "CopyNum" & i Then
                VarsExist(i) = True
            End If
        Next
    Next
    'Init variables if they don't exist
    For i = 1 To 4
        If Not VarsExist(i) Then
            ActiveDocument.Variables.Add "CopyNum" & i, 0
        End If
    Next

    'Ask how many to print
    NumCopiesToPrint = CLng(InputBox("How many pages?", "Number and print", "1"))

    'Ask where to start numbering
    StartAt = CLng(InputBox("At which page should be started?", "Number and print", CStr(ActiveDocument.Variables("CopyNum4") + 1)))

    'Loop through print-write-print cycle
    For Counter = 0 To NumCopiesToPrint - 1
        'Update the document variables
        For i = 1 To 4
            ActiveDocument.Variables("CopyNum" & i) = StartAt + Counter * 4 + i - 1
        Next
        'Update Fields
        ActiveDocument.Fields.Update
        If ActiveDocument.Shapes.Count > 0 Then
            Dim Shape As Variant
            For Each Shape In ActiveDocument.Shapes
                If Shape.TextFrame.HasText Then
                    Shape.TextFrame.TextRange.Fields.Update
                End If
            Next
        End If
        'Print numbered copy
        ActiveDocument.PrintOut PrintZoomRow:=2, PrintZoomColumn:=2
    Next Counter
    Cancel = True
End Sub
