VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Event definition class. The actual functionality on what happens when the event is fired is implemented here."
'@Folder "Continual numbering"
'@ModuleDescription "Event definition class. The actual functionality on what happens when the event is fired is implemented here."
'@IgnoreModule EncapsulatePublicField
'@MemberAttribute VB_VarHelpID, -1
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1
Option Explicit

'Constants
Private Const InputTitleText As String = "Number and print"
Private Const NumPagesQuestion As String = "How many pages?"
Private Const StartPagesQuestion As String = "At which page should be started?"
Private Const PagesPerRow As Long = 2
Private Const PagesPerColumn As Long = 2
Private Const PagesPerSheet As Long = PagesPerRow * PagesPerColumn
Private Const CopyNumberLabel As String = "CopyNum"


'@Description "Before the document is printed, ask for input and change document variables to achieve continuously numbered pages."
Private Sub App_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)
Attribute App_DocumentBeforePrint.VB_Description = "Before the document is printed, ask for input and change document variables to achieve continuously numbered pages."
    'Variables
    Dim i As Long
    Dim NumCopiesToPrint As Long
    Dim Counter As Long
    Dim StartAt As Long

    'Ensure variables exists
    Dim VarsExist(1 To PagesPerSheet) As Boolean
    For i = 1 To PagesPerSheet
        VarsExist(i) = False
    Next
    Dim Var As Variable
    For Each Var In ActiveDocument.Variables
        For i = 1 To PagesPerSheet
            If Var.Name = CopyNumberLabel & i Then
                VarsExist(i) = True
            End If
        Next
    Next
    'Init variables if they don't exist
    For i = 1 To 4
        If Not VarsExist(i) Then
            ActiveDocument.Variables.Add CopyNumberLabel & i, 0
        End If
    Next

    'Ask how many to print
    NumCopiesToPrint = CLng(InputBox(NumPagesQuestion, InputTitleText, "1"))
    'Ask where to start numbering
    StartAt = CLng(InputBox(StartPagesQuestion, InputTitleText, CStr(ActiveDocument.Variables.Item(CopyNumberLabel & PagesPerSheet).Value + 1)))

    'Loop through print-write-print cycle
    For Counter = 0 To NumCopiesToPrint - 1
        'Update the document variables
        For i = 1 To PagesPerSheet
            ActiveDocument.Variables.Item(CopyNumberLabel & i).Value = StartAt + Counter * PagesPerSheet + i - 1
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
        ActiveDocument.PrintOut PrintZoomRow:=PagesPerRow, PrintZoomColumn:=PagesPerColumn
    Next Counter
    Cancel = True
End Sub
