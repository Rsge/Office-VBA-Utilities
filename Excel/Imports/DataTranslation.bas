Attribute VB_Name = "DataTranslation"
Attribute VB_Description = "Translates data between languages using a pre-defined dictionary."
'@Folder("Imports")
'@ModuleDescription("Translates data between languages using a pre-defined dictionary.")
Option Explicit

' String constants
'@VariableDescription("The first row with data.")
Private Const m_startingRow As Long = 2
Attribute m_startingRow.VB_VarDescription = "The first row with data."
'@VariableDescription("Remove this string from front of packing unit.")
Private Const m_removeThis As String = "-"
Attribute m_removeThis.VB_VarDescription = "Remove this string from front of packing unit."
'@VariableDescription("Keep packing unit before this string, cut rest.")
Private Const m_keepBeforeThis As String = "+"
Attribute m_keepBeforeThis.VB_VarDescription = "Keep packing unit before this string, cut rest."

' Integer constants
'@VariableDescription("The column containing the item number.")
Private Const m_itemColumn As Long = 1
Attribute m_itemColumn.VB_VarDescription = "The column containing the item number."
'@VariableDescription("The column containing english localization.")
Private Const m_englishColumn As Long = 4
Attribute m_englishColumn.VB_VarDescription = "The column containing english localization."
'@VariableDescription("The column containing german localization.")
Private Const m_germanColumn As Long = 3
Attribute m_germanColumn.VB_VarDescription = "The column containing german localization."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Translates data between languages in different columns.")
'@ExcelHotkey T
Public Sub Translate()
Attribute Translate.VB_Description = "Translates data between languages in different columns."
Attribute Translate.VB_ProcData.VB_Invoke_Func = "T\n14"
    ' Variables
    Dim dictEnDeP As Object
    Set dictEnDeP = CreateObject("Scripting.Dictionary")
    Dim dictDeEnP As Object
    Set dictDeEnP = CreateObject("Scripting.Dictionary")
    Dim dictEnDePU As Object
    Set dictEnDePU = CreateObject("Scripting.Dictionary")
    Dim dictDeEnPU As Object
    Set dictDeEnPU = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = m_startingRow

    ' Add translations.
    Dim key As Variant
    dictEnDeP.Add "Example", "Beispiel"
    For Each key In dictEnDeP.Keys()
        dictDeEnP.Add dictEnDeP.Item(key), key
    Next
    dictEnDePU.Add "Fortytwo", "Zweiundvierzig"
    For Each key In dictEnDePU.Keys()
        dictDeEnPU.Add dictEnDePU.Item(key), key
    Next
    ' Translate data in Excel table.
    Dim englishCell As Range
    Dim germanCell As Range
    Do Until LenB(ActiveSheet.Cells.Item(i, m_itemColumn).Value) = 0
        Set englishCell = ActiveSheet.Cells.Item(i, m_englishColumn)
        Set germanCell = ActiveSheet.Cells.Item(i, m_germanColumn)
        If LenB(englishCell.Value) = 0 And LenB(germanCell.Value) > 0 Then
            englishCell.Value = GetPackagingTranslation(germanCell.Value, dictDeEnP, dictDeEnPU)
        ElseIf LenB(germanCell.Value) = 0 And LenB(englishCell.Value) > 0 Then
            germanCell.Value = GetPackagingTranslation(englishCell.Value, dictEnDeP, dictEnDePU)
        End If
        i = i + 1
    Loop
End Sub

'@Description("Translates string using dictionaries.")
Private Function GetPackagingTranslation(ByVal str As String, ByVal dictP As Object, ByVal dictPu As Object) As String
Attribute GetPackagingTranslation.VB_Description = "Translates string using dictionaries."
    Dim packaging As String
    Dim translation As String
    packaging = Replace(Left$(str, InStr(str, m_keepBeforeThis) - 1), m_removeThis, vbNullString)
    translation = Replace(str, packaging, dictP.Item(packaging))
    Dim packingUnit As String
    Dim trans As Variant
    packingUnit = Right$(str, Len(str) - InStr(str, m_keepBeforeThis) - Len(m_keepBeforeThis))
    For Each trans In dictPu.Keys()
        If InStr(packingUnit, trans) > 0 Then
            translation = Replace(translation, packingUnit, Replace(packingUnit, trans, dictPu.Item(trans)))
            Exit For
        End If
    Next
    GetPackagingTranslation = translation
End Function
