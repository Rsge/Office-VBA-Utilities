Attribute VB_Name = "Dict"
Attribute VB_Description = "Creation of dictionaries from entries in Excel tables."
'@Folder("DocumentCreation.Base")
'@ModuleDescription("Creation of dictionaries from entries in Excel tables.")
Option Explicit
'Option Private Module

'TODO Solve this w/ Class &| Types

'Dicts
Private m_dictDocNative As Object
Private m_dictDocTransl As Object
Private m_dictDecimalsNative As Object
Private m_dictDecimalsTransl As Object
Private m_dictData As Object
Private m_listDocSkip As Object
'@Description("Dictionary with mapping of native language document categories to respective data categories.")
Public Property Get DictDocNative() As Object
Attribute DictDocNative.VB_Description = "Dictionary with mapping of native language document categories to respective data categories."
    Set DictDocNative = m_dictDocNative
End Property
'@Description("Dictionary with mapping of translated document categories to respective data categories.")
Public Property Get DictDocTransl() As Object
Attribute DictDocTransl.VB_Description = "Dictionary with mapping of translated document categories to respective data categories."
    Set DictDocTransl = m_dictDocTransl
End Property
'@Description("Dictionary with mapping of decimal and padding places to respective native data categories.")
Public Property Get DictDecimalsNative() As Object
Attribute DictDecimalsNative.VB_Description = "Dictionary with mapping of decimal and padding places to respective native data categories."
    Set DictDecimalsNative = m_dictDecimalsNative
End Property
'@Description("Dictionary with mapping of decimal and padding places to respective translated data categories.")
Public Property Get DictDecimalsTransl() As Object
Attribute DictDecimalsTransl.VB_Description = "Dictionary with mapping of decimal and padding places to respective translated data categories."
    Set DictDecimalsTransl = m_dictDecimalsTransl
End Property
'@Description("Dictionary with mapping of native language data categories to respective document categories.")
Public Property Get DictData() As Object
Attribute DictData.VB_Description = "Dictionary with mapping of native language data categories to respective document categories."
    Set DictData = m_dictData
End Property
'@Description("List of entries in document categories region to skip over.")
Public Property Get ListDocSkip() As Object
Attribute ListDocSkip.VB_Description = "List of entries in document categories region to skip over."
    Set ListDocSkip = m_listDocSkip
End Property

' ————————————————————————————————————————————————————— '

'@Description("Determines if one string that must and at least one that can be given is.")
Private Function HasEntries(ByVal must As String, ByVal firstCan As String, ByVal secondCan As String) As Boolean
Attribute HasEntries.VB_Description = "Determines if one string that must and at least one that can be given is."
    HasEntries = Not (IsEmpty(must) Or (IsEmpty(firstCan) And IsEmpty(secondCan)))
End Function

'@Description("Inits the dictionary.")
Public Sub InitDicts()
Attribute InitDicts.VB_Description = "Inits the dictionary."
    ' Create dicts.
    Set m_dictDocNative = CreateObject("Scripting.Dictionary")
    Set m_dictDocTransl = CreateObject("Scripting.Dictionary")
    Set m_dictDecimalsNative = CreateObject("Scripting.Dictionary")
    Set m_dictDecimalsTransl = CreateObject("Scripting.Dictionary")
    Set m_dictData = CreateObject("Scripting.Dictionary")
    Set m_listDocSkip = CreateObject("System.Collections.ArrayList")
    ' Declarations
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Item(MappingSheetName)
    Dim docToDataTable As Range
    Set docToDataTable = ws.Range(DocToDataTableName)
    Dim dataToDocTable As Range
    Set dataToDocTable = ws.Range(DataToDocTableName)
    Dim docSkipTable As Range
    Set docSkipTable = ws.Range(DocSkipTableName)
    Dim row_ As Variant
    Dim doc(1) As String
    Dim data As String
    Dim decimalsStr As String
    Dim decimals As Long
    
    ' Fill doc to data dicts.
    For Each row_ In docToDataTable.Rows
        doc(0) = GetCellValueR(row_, DocToDataNativeColum)
        doc(1) = GetCellValueR(row_, DocToDataTranslColum)
        data = GetCellValueR(row_, DocToDataDataColumn)
        decimalsStr = GetCellValueR(row_, DocToDataDecimalColumn)
        If HasEntries(data, doc(0), doc(1)) Then
            On Error GoTo DictErr
            If Not IsEmpty(doc(0)) Then m_dictDocNative.Add doc(0), data
            If Not IsEmpty(doc(1)) Then m_dictDocTransl.Add doc(1), data
            On Error GoTo DecimalErr
            If Not IsEmpty(decimalsStr) Then
                decimals = CLng(decimalsStr)
                If Not IsEmpty(doc(0)) Then m_dictDecimalsNative.Add doc(0), decimals
                If Not IsEmpty(doc(1)) Then m_dictDecimalsTransl.Add doc(1), decimals
            End If
            On Error GoTo 0
        End If
    Next
    ' Fill data to doc dict.
    For Each row_ In dataToDocTable.Rows
        data = GetCellValueR(row_, DataToDocDataColumn)
        doc(0) = GetCellValueR(row_, DataToDocNativeColum)
        doc(1) = GetCellValueR(row_, DataToDocTranslColum)
        If HasEntries(data, doc(0), doc(1)) Then
            On Error GoTo DictErr
            m_dictData.Add data, doc
            On Error GoTo 0
        End If
    Next
    ' Fill doc skip list.
    Dim cell As Variant
    For Each cell In docSkipTable.Cells
        If Not IsEmpty(cell.Value) Then
            If Not m_listDocSkip.Contains(cell.Value) Then
                m_listDocSkip.Add cell.Value
            Else
                GoTo ListErr
            End If
        End If
    Next
    Exit Sub
    
    ' Raise error with useful error description.
DictErr:
    Err.Raise vbObjectError + 1, "Dict", FormatString(DuplicateCategoryIDError, doc(0) & RowSep & doc(1) & RowSep & data)
    Exit Sub
DecimalErr:
    Err.Raise vbObjectError + 2, "Decimal", FormatString(DecimalEntryError, decimalsStr)
    Exit Sub
ListErr:
    Err.Raise vbObjectError + 3, "List", FormatString(DuplicateCategoryIDError, cell.Value)
End Sub
