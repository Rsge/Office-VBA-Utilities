Attribute VB_Name = "DataTranslation"
Attribute VB_Description = "Translates data between languages using a pre-defined dictionary."
'@Folder("Imports")
'@ModuleDescription("Translates data between languages using a pre-defined dictionary.")
Option Explicit

' String constants
'@VariableDescription("Remove this string from front of packing unit.")
Private Const m_removeThis As String = "xx "
Attribute m_removeThis.VB_VarDescription = "Remove this string from front of packing unit."
'@VariableDescription("Keep packing unit before this string, cut rest.")
Private Const m_keepBeforeThis As String = "yy"
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
    Dim dictDeEn1 As Object
    Set dictDeEn1 = CreateObject("Scripting.Dictionary")
    Dim dictEnDe1 As Object
    Set dictEnDe1 = CreateObject("Scripting.Dictionary")
    Dim dictDeEn2 As Object
    Set dictDeEn2 = CreateObject("Scripting.Dictionary")
    Dim dictEnDe2 As Object
    Set dictEnDe2 = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 2

    ' Add translations.
    Dim key As Variant
    dictDeEn1.Add "Example", "Beispiel"
    For Each key In dictDeEn1.keys()
        dictEnDe1.Add dictDeEn1.Item(key), key
    Next
    dictDeEn2.Add "Fortytwo", "Zweiundvierzig"
    For Each key In dictDeEn2.keys()
        dictEnDe2.Add dictDeEn2.Item(key), key
    Next
    ' Translate data in Excel table.
    Dim german As String
    Dim english As String
    Do While LenB(ActiveSheet.Cells(i, m_itemColumn)) <> 0
        english = ActiveSheet.Cells(i, m_englishColumn)
        german = ActiveSheet.Cells(i, m_germanColumn)
        If LenB(english) > 0 And LenB(german) = 0 Then
            ActiveSheet.Cells(i, m_germanColumn) = GetPackagingTranslation(english, dictEnDe1, dictEnDe2)
        ElseIf LenB(german) > 0 And LenB(english) = 0 Then
            ActiveSheet.Cells(i, m_englishColumn) = GetPackagingTranslation(german, dictDeEn1, dictDeEn2)
        End If
        i = i + 1
    Loop
End Sub

'@Description("Translates string using dictionaries.")
Private Function GetPackagingTranslation(ByVal str As String, ByVal dict1 As Object, ByVal dict2 As Object) As String
Attribute GetPackagingTranslation.VB_Description = "Translates string using dictionaries."
    Dim packaging As String
    Dim translation As String
    packaging = Replace(Left$(str, InStr(str, m_keepBeforeThis) - 1), m_removeThis, vbNullString)
    translation = Replace(str, packaging, dict1.Item(packaging))
    Dim packingUnit As String
    Dim trans As Variant
    packingUnit = Right$(str, Len(str) - InStr(str, m_keepBeforeThis) - Len(m_keepBeforeThis))
    For Each trans In dict2.keys()
        If InStrB(packingUnit, trans) > 0 Then
            translation = Replace(translation, packingUnit, Replace(packingUnit, trans, dict2.Item(trans)))
            Exit For
        End If
    Next
    GetPackagingTranslation = translation
End Function
