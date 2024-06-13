Attribute VB_Name = "CategoryUnification"
Attribute VB_Description = "Handles unification of category names in data files."
'@IgnoreModule UnreachableCase
'@Folder("DocumentCreation.Unification")
'@ModuleDescription("Handles unification of category names in data files.")
Option Explicit

' Runtime constants
' Change each run to get results you want.
'@VariableDescription("List only mode (no changes done to data).")
Private Const m_listOnly As Boolean = True
Attribute m_listOnly.VB_VarDescription = "List only mode (no changes done to data)."
'@VariableDescription("Mode for looking at the next lines after a match. 0 = Off, 1 = Replace in first, 2 = Replace in second, 3 = Replace in third.")
Private Const m_lookAtNextMode As Long = 0
Attribute m_lookAtNextMode.VB_VarDescription = "Mode for looking at the next lines after a match. 0 = Off, 1 = Replace in first, 2 = Replace in second, 3 = Replace in third."
'@VariableDescription("Pattern of regex for first cell. Used to unify column info.")
Private Const m_regexPattern As String = "x"
Attribute m_regexPattern.VB_VarDescription = "Pattern of regex for first cell. Used to unify column info."
'@VariableDescription("Pattern of regex for the next cell. Used to unify column info.")
Private Const m_regexPatternNext As String = "-"
Attribute m_regexPatternNext.VB_VarDescription = "Pattern of regex for the next cell. Used to unify column info."
'@VariableDescription("What to replace the found match with.")
Private Const m_regexReplace As String = "y"
Attribute m_regexReplace.VB_VarDescription = "What to replace the found match with."

' Path constants
'@VariableDescription("Name of file for saving used column names, new per file.")
Private Const m_listFileNameFileAdds As String = "Categories - File.txt"
Attribute m_listFileNameFileAdds.VB_VarDescription = "Name of file for saving used column names, new per file."
'@VariableDescription("Name of file for saving used column names in alphabetical order with each line as it's own category.")
Private Const m_listFileNameSeparateAlphabetical As String = "Categories - Single, Alphab.txt"
Attribute m_listFileNameSeparateAlphabetical.VB_VarDescription = "Name of file for saving used column names in alphabetical order with each line as it's own category."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Lists all categories found in data and uses regex to unify category names.")
Public Sub ListAndUnifiyCategories()
Attribute ListAndUnifiyCategories.VB_Description = "Lists all categories found in data and uses regex to unify category names."
    ' Declarations
    Dim dataWBFileName As String
    dataWBFileName = Dir(DataWBsPath & AllDataWBFilesPattern)
    Dim ignore As Variant
    Dim ignores() As String
    ignores = Split(IgnoreList, ListSep)
    Dim dataWB As Workbook
    Dim dataWS As Worksheet
    Dim changed As Boolean
    changed = False
    Dim lastColumn As Long
    Dim categoryCounts As Object
    Set categoryCounts = CreateObject("Scripting.Dictionary")
    Dim categoriesEdited As Object
    Set categoriesEdited = CreateObject("System.Collections.ArrayList")
    Dim categoriesFileNew As Object
    Set categoriesFileNew = CreateObject("System.Collections.ArrayList")
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = m_regexPattern
    Dim regexNext As Object
    Set regexNext = CreateObject("VBScript.RegExp")
    regexNext.Pattern = m_regexPatternNext
    Dim currentCell As Range
    Dim nextCell As Range
    Dim categoryName As String
    Dim joinedCategoryName As String
    Dim currentColumn As Long
    Dim currentRow As Long
    Dim first As Boolean
    first = True
    Dim replaced As String
    Dim matches As Boolean
    
    ' Go through all files fitting the pattern.
    Do Until IsEmpty(dataWBFileName)
        ' If file name contains ignore marker symbol, skip it, continuing with next file.
        For Each ignore In ignores
            If Contains(dataWBFileName, CStr(ignore)) _
            Then GoTo Skip
        Next
        ' Open current workbook and look at first sheet.
        Set dataWB = Workbooks.Open(DataWBsPath & dataWBFileName)
        Set dataWS = dataWB.Sheets.[_Default](1)
        ' If first sheet contains the first ignore symbol, move second sheet before first and look at that sheet.
        changed = ChooseCorrectSheet(dataWB, dataWS, ignores(0))
        ' Go through all used columns.
        currentColumn = DataStartingColumn
        lastColumn = GetLastColumnIndex(dataWS, DataCategoryStartingRow)
        Do While currentColumn <= lastColumn
            ' Edit files if not in list only mode, otherwise print the changed string to debug.
            Set currentCell = GetCell(dataWS, DataCategoryStartingRow, currentColumn)
            Select Case m_lookAtNextMode
                ' Just look at and replace in current cell.
                Case 0
                    currentRow = DataCategoryStartingRow
                    Do While currentRow <= DataCategoryStoppingRow
                        replaced = regex.Replace(currentCell.Value, m_regexReplace)
                        If replaced <> currentCell.Value Then
                            If m_listOnly Then
                                If Not categoriesEdited.Contains(replaced) Then
                                    categoriesEdited.Add replaced
                                    Debug.Print replaced
                                End If
                            Else
                                currentCell.Value = replaced
                                changed = True
                            End If
                        End If
                        currentRow = currentRow + 1
                        Set currentCell = GetCell(dataWS, currentRow, currentColumn)
                    Loop
                ' Look at current and next cell and replace in current.
                Case 1
                    replaced = regex.Replace(currentCell.Value, m_regexReplace)
                    Set nextCell = GetCell(dataWS, DataCategoryStartingRow + 1, currentColumn)
                    matches = regexNext.Test(nextCell.Value)
                    If replaced <> currentCell.Value And matches Then
                        If m_listOnly Then
                            categoryName = replaced & RowSep & nextCell.Value
                            If Not categoriesEdited.Contains(categoryName) Then
                                categoriesEdited.Add categoryName
                                Debug.Print categoryName
                            End If
                        Else
                            currentCell.Value = replaced
                            changed = True
                        End If
                    End If
                ' Look at current and next cell and replace in next.
                Case 2
                    matches = regex.Test(currentCell.Value)
                    Set nextCell = GetCell(dataWS, DataCategoryStartingRow + 1, currentColumn)
                    replaced = regexNext.Replace(nextCell.Value, m_regexReplace)
                    If matches And replaced <> nextCell.Value Then
                        If m_listOnly Then
                            categoryName = replaced & RowSep & nextCell.Value
                            If Not categoriesEdited.Contains(categoryName) Then
                                categoriesEdited.Add categoryName
                                Debug.Print categoryName
                            End If
                        Else
                            nextCell.Value = replaced
                            changed = True
                        End If
                    End If
                ' Look at current cell and cell after next and replace in cell after next.
                Case 3
                    matches = regex.Test(currentCell.Value)
                    Set nextCell = GetCell(dataWS, DataCategoryStartingRow + 2, currentColumn)
                    replaced = regexNext.Replace(nextCell.Value, m_regexReplace)
                    If matches And replaced <> nextCell.Value Then
                        If m_listOnly Then
                            categoryName = currentCell.Value & RowSep & replaced
                            If Not categoriesEdited.Contains(categoryName) Then
                                categoriesEdited.Add categoryName
                                Debug.Print categoryName
                            End If
                        Else
                            nextCell.Value = replaced
                            changed = True
                        End If
                    End If
            End Select
            currentRow = DataCategoryStartingRow
            joinedCategoryName = vbNullString
            Do While currentRow <= DataCategoryStoppingRow
                ' Get category name.
                categoryName = GetCellValue(dataWS, currentRow, currentColumn)
                If Not categoryCounts.Exists(categoryName) Then
                    categoryCounts.Add categoryName, 1
                Else
                    categoryCounts.Item(categoryName) = categoryCounts.Item(categoryName) + 1
                End If
                joinedCategoryName = joinedCategoryName & categoryName & RowSep
                currentRow = currentRow + 1
            Loop
            ' Join category names.
            joinedCategoryName = Space$(4) & RemoveLast(joinedCategoryName, Len(RowSep))
            If Not categoriesFileNew.Contains(joinedCategoryName) Then
                If first Then
                    categoriesFileNew.Add dataWBFileName
                    first = False
                End If
                categoriesFileNew.Add joinedCategoryName
            End If
            currentColumn = currentColumn + 1
        Loop
        ' Finish processing.
        first = True
        If changed Then
            dataWB.Save
            changed = False
        End If
        dataWB.Close
Skip:
        dataWBFileName = Dir
        DoEvents
    Loop
    
    ' Sort categories.
    Dim categories As Object
    Set categories = CreateObject("System.Collections.ArrayList")
    Dim entry As Variant
    For Each entry In categoryCounts.Keys
        categories.Add entry
    Next
    categories.Sort
    ' Save results in txt.
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Open ListFilesPath & m_listFileNameSeparateAlphabetical For Output As fileNumber
    For Each entry In categories
        Print #fileNumber, categoryCounts.Item(entry) & vbTab & entry
    Next
    Close fileNumber
    Open ListFilesPath & m_listFileNameFileAdds For Output As fileNumber
    For Each entry In categoriesFileNew
        Print #fileNumber, entry
    Next
    Close fileNumber
End Sub
