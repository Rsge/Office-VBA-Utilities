VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrintEventMultiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Event definition. Actual functionality of the event is implemented here."
'@Folder("SequentiallyNumberedCopies")
'@ModuleDescription("Event definition. Actual functionality of the event is implemented here.")
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
'@Ignore EncapsulatePublicField
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

' Constants
'@VariableDescription("Input window's title text.")
Private Const m_inputTitleText As String = "Number and print"
Attribute m_inputTitleText.VB_VarDescription = "Input window's title text."
'@VariableDescription("Question for how many sheets should be printed.")
Private Const m_numSheetsQuestion As String = "How many sheets?"
Attribute m_numSheetsQuestion.VB_VarDescription = "Question for how many sheets should be printed."
'@VariableDescription("Question at which page to start.")
Private Const m_startPagesQuestion As String = "At which page number should be started?"
Attribute m_startPagesQuestion.VB_VarDescription = "Question at which page to start."
'@VariableDescription("How many pages there are in a row on a sheet.")
Private Const m_pagesPerRow As Long = 2
Attribute m_pagesPerRow.VB_VarDescription = "How many pages there are in a row on a sheet."
'@VariableDescription("How many pages there are in a column on a sheet.")
Private Const m_pagesPerColumn As Long = 2
Attribute m_pagesPerColumn.VB_VarDescription = "How many pages there are in a column on a sheet."
'@VariableDescription("How many pages there are on a sheet, total.")
Private Const m_pagesPerSheet As Long = m_pagesPerRow * m_pagesPerColumn
Attribute m_pagesPerSheet.VB_VarDescription = "How many pages there are on a sheet, total."
'@VariableDescription("Copy number variable name.")
Private Const m_copyNumberLabel As String = "CopyNum"
Attribute m_copyNumberLabel.VB_VarDescription = "Copy number variable name."


'@Description("Before the document is printed, ask for input and change document variables to achieve continuously numbered pages.")
Private Sub App_DocumentBeforePrint(ByVal doc As Document, cancel As Boolean)
Attribute App_DocumentBeforePrint.VB_Description = "Before the document is printed, ask for input and change document variables to achieve continuously numbered pages."
    ' Variables
    Dim i As Long
    Dim numCopiesToPrint As Long
    Dim counter As Long
    Dim startAt As Long

    ' Ensure variables exists.
    Dim varsExist(1 To m_pagesPerSheet) As Boolean
    For i = 1 To m_pagesPerSheet
        varsExist(i) = False
    Next
    Dim var As Variable
    For Each var In ActiveDocument.Variables
        For i = 1 To m_pagesPerSheet
            If var.Name = m_copyNumberLabel & i Then
                varsExist(i) = True
            End If
        Next
    Next
    ' Init variables if they don't exist.
    For i = 1 To m_pagesPerSheet
        If Not varsExist(i) Then
            ActiveDocument.Variables.Add m_copyNumberLabel & i, 0
        End If
    Next

    On Error GoTo Abort
    ' Ask how many to print.
    numCopiesToPrint = CLng(InputBox(m_numSheetsQuestion, m_inputTitleText, "1"))
    ' Ask where to start numbering.
    startAt = CLng(InputBox(m_startPagesQuestion, m_inputTitleText, CStr(ActiveDocument.Variables.Item(m_copyNumberLabel & m_pagesPerSheet).Value + 1)))
    On Error GoTo 0

    ' Loop through print-write-print cycle.
    For counter = 0 To numCopiesToPrint - 1
        ' Update the document variables.
        For i = 1 To m_pagesPerSheet
            ActiveDocument.Variables.Item(m_copyNumberLabel & i).Value = Format$(startAt + counter * m_pagesPerSheet + i - 1, "00#")
        Next
        ' Update Fields.
        ActiveDocument.Fields.Update
        If ActiveDocument.Shapes.Count > 0 Then
            Dim Shape As Variant
            For Each Shape In ActiveDocument.Shapes
                If Shape.TextFrame.HasText Then
                    Shape.TextFrame.TextRange.Fields.Update
                End If
            Next
        End If
        ' Print numbered copy.
        ActiveDocument.PrintOut PrintZoomRow:=m_pagesPerRow, PrintZoomColumn:=m_pagesPerColumn
    Next
Abort:
    cancel = True
End Sub
