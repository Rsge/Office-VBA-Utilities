Attribute VB_Name = "HotkeyActions"
Attribute VB_Description = "Adds various useful hotkeys for different often-used actions."
'@Folder("TextEditing")
'@ModuleDescription("Sets currently selected text to monospace font.")
Option Explicit

'@ExcelHotkey Alt+M
Public Sub ToggleMonospace()
Attribute ToggleMonospace.VB_ProcData.VB_Invoke_Func = "A\n14"
    With Selection.Font
        If .Name <> "Arial" Then
            .Name = "Arial"
        Else
            .Name = "Consolas"
        End If
    End With
End Sub

'@ExcelHotkey Alt+L
Public Sub ToggleLanguage()
Attribute ToggleLanguage.VB_ProcData.VB_Invoke_Func = "A\n14"
    With Selection
        Select Case .LanguageID
            Case wdGerman
                .LanguageID = wdEnglishUS
            Case wdEnglishUS
                .LanguageID = wdGerman
        End Select
    End With
End Sub

'@ExcelHotkey Alt+D
Public Sub InsertDate()
Attribute InsertDate.VB_ProcData.VB_Invoke_Func = "A\n14"
    Selection.TypeText Format$(Date, "yyyy-MM-dd")
End Sub
