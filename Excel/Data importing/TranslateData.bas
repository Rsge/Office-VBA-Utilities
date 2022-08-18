Attribute VB_Name = "TranslateData"
Attribute VB_Description = "Importing string data from text file to excel table."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Import"
'@ModuleDescription "Importing string data from text file to excel table."
Option Explicit

'String constants
'@VariableDescription "Remove this string from front of packing unit."
Private Const m_removeThis As String = "xx"
Attribute m_removeThis.VB_VarDescription = "Remove this string from front of packing unit."
'@VariableDescription "Keep packing unit before this string, cut rest."
Private Const m_keepBeforeThis As String = "yy"
Attribute m_keepBeforeThis.VB_VarDescription = "Keep packing unit before this string, cut rest."

'Int constants
'@VariableDescription "Column
Private Const m_itemColumn As Long = 1
Private Const m_germanColumn As Long = 3
Private Const m_englishColumn As Long = 4


'@EntryPoint
'@Description "Translates data between languages in different columns."
Public Sub Translate()
Attribute Translate.VB_Description = "Translates data between languages in different columns."
    'Variables
    Dim dictEnDeP1 As Object
    Set dictEnDeP1 = CreateObject("Scripting.Dictionary")
    Dim dictDeEnP1 As Object
    Set dictDeEnP1 = CreateObject("Scripting.Dictionary")
    Dim dictDeEnP2 As Object
    Set dictDeEnP2 = CreateObject("Scripting.Dictionary")
    Dim dictDeEnP2 As Object
    Set dictDeEnP2 = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 2
    
    'Adding translations
    Dim key As Variant
    dictEnDeP1.Add "Example", "Beispiel"
    For Each key In dictEnDeP1.keys()
        dictDeEnP1.Add dictEnDeP1(key), key
    Next
    dictDeEnP2.Add "Fortytwo", "Zweiundvierzig"
    For Each key In dictDeEnP2.keys()
        dictDeEnP2.Add dictDeEnP2(key), key
    Next
    'Translating data in Excel table
    Dim german As String
    Dim english As String
    Do While LenB(ActiveSheet.Cells(i, m_itemColumn)) <> 0
        german = ActiveSheet.Cells(i, m_germanColumn)
        english = ActiveSheet.Cells(i, m_englishColumn)
        If LenB(english) > 0 And LenB(german) = 0 Then
            ActiveSheet.Cells(i, m_germanColumn) = GetPackagingTranslation(english, dictEnDeP1, dictDeEnP2)
        ElseIf LenB(german) > 0 And LenB(english) = 0 Then
            ActiveSheet.Cells(i, m_englishColumn) = GetPackagingTranslation(german, dictDeEnP1, dictDeEnP2)
        End If
        i = i + 1
    Loop
End Sub

'@Description "Translates string using dictionaries."
Private Function GetPackagingTranslation(ByVal str As String, ByVal dictP1 As Object, ByVal dictP2 As Object) As String
Attribute Translate.VB_Description = "Translates string using dictionaries."
    Dim packaging As String
    Dim translation As String
    packaging = Replace(Left$(str, InStr(str, m_keepBeforeThis) - 1), m_removeThis, vbNullString)
    translation = Replace(str, packaging, dictP1(packaging))
    Dim packingUnit As String
    Dim trans As Variant
    packingUnit = Right$(str, Len(str) - InStr(str, m_keepBeforeThis) - Len(m_keepBeforeThis))
    For Each trans In dictP2.keys()
        If InStrB(packingUnit, trans) > 0 Then
            translation = Replace(translation, packingUnit, Replace(packingUnit, trans, dictP2(trans)))
            Exit For
        End If
    Next
    GetPackagingTranslation = translation
End Function
