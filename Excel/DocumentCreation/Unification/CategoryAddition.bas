Attribute VB_Name = "CategoryAddition"
Attribute VB_Description = "Handles adding a new category to data files."
'@IgnoreModule UnreachableCase
'@Folder("DocumentCreation.Unification")
'@ModuleDescription("Handles adding a new category to data files.")
Option Explicit

' Runtime constants
' Change each run to get results you want.
'@VariableDescription("List only mode (no changes done to data).")
Private Const m_listOnly As Boolean = True
Attribute m_listOnly.VB_VarDescription = "List only mode (no changes done to data)."
'@VariableDescription("How many columns left and right of found cell should be listed.")
Private Const m_listBreadth As Long = 1
Attribute m_listBreadth.VB_VarDescription = "How many columns left and right of found cell should be listed."
'@VariableDescription("Mode for where to insert a column. 0 = Place before, 1 = Place after.")
Private Const m_insertMode As Long = 0
Attribute m_insertMode.VB_VarDescription = "Mode for where to insert a column. 0 = Place before, 1 = Place after."
'@VariableDescription("Pattern of regex for first row's cell to find.")
Private Const m_regexPattern As String = "x"
Attribute m_regexPattern.VB_VarDescription = "Pattern of regex for first row's cell to find."
'@VariableDescription("Pattern of regex for the neighboring cells to check if insert categegory is already there. Set to vbNullString to disable check.")
Private Const m_regexPatternNeighbors As String = "-"
Attribute m_regexPatternNeighbors.VB_VarDescription = "Pattern of regex for the neighboring cells to check if insert categegory is already there. Set to vbNullString to disable check."
'@VariableDescription("What category to insert. Rows separated by row separator. Only the amount of category rows is used.")
Private Const m_insertCategory As String = "x" & RowSep _
                                         & vbNullString & RowSep _
                                         & "y"
Attribute m_insertCategory.VB_VarDescription = "What category to insert. Rows separated by row separator. Only the amount of category rows is used."

' Path constants
'@VariableDescription("Name of file for saving used column names, new per file.")
Private Const m_listFileNameCategoryAdditions As String = "Categories - Added.txt"
Attribute m_listFileNameCategoryAdditions.VB_VarDescription = "Name of file for saving used column names, new per file."

' ————————————————————————————————————————————————————— '


'@EntryPoint
'@Description("Lists specific categories found in data, adds a new one near it or if that already exists moves it to the right place.")
Public Sub ListSpecificAndAddNewCategories()
Attribute ListSpecificAndAddNewCategories.VB_Description = "Lists specific categories found in data, adds a new one near it or if that already exists moves it to the right place."
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
    Dim currentRow As Long
    Dim currentColumn As Long
    Dim currentCell As Range
    Dim lastColumn As Long
    Dim foundColumn As Long
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = m_regexPattern
    Dim regexNeighbors As Object
    Set regexNeighbors = CreateObject("VBScript.RegExp")
    regexNeighbors.Pattern = m_regexPatternNeighbors
    Dim joinedCategoryName As String
    Dim categoriesFile As Object
    Set categoriesFile = CreateObject("System.Collections.ArrayList")
    Dim insertRows() As String
    Dim prevValue As String
    Dim nextValue As String
    Dim prevMatches As Boolean
    Dim nextMatches As Boolean
    Dim i As Long
    
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
            ' Edit files only if not in list only mode.
            Set currentCell = GetCell(dataWS, DataCategoryStartingRow, currentColumn)
            If regex.Test(currentCell.Value) Then
                foundColumn = currentColumn
                i = 0
                insertRows = Split(m_insertCategory, RowSep)
                If Not m_listOnly Then
                    prevValue = GetCellValue(dataWS, DataCategoryStartingRow, currentColumn - 1)
                    nextValue = GetCellValue(dataWS, DataCategoryStartingRow, currentColumn + 1)
                    prevMatches = regexNeighbors.Test(prevValue)
                    nextMatches = regexNeighbors.Test(nextValue)
                    With dataWS.Columns
                        If Not IsEmpty(regexNeighbors.Pattern) _
                        And (prevMatches Or nextMatches) Then
                                Select Case m_insertMode
                                    ' Insert column before match.
                                    Case 0
                                        If nextMatches Then
                                            .Item(currentColumn).Insert
                                            .Item(currentColumn).Value = .Item(currentColumn + 2).Value
                                            .Item(currentColumn + 2).Delete
                                            foundColumn = foundColumn + 1
                                        Else
                                            currentColumn = currentColumn - 1
                                        End If
                                    ' Insert column after match.
                                    Case 1
                                        If prevMatches Then
                                            .Item(currentColumn + 1).Insert
                                            .Item(currentColumn + 1).Value = .Item(currentColumn - 1).Value
                                            .Item(currentColumn - 1).Delete
                                            foundColumn = foundColumn - 1
                                        Else
                                            currentColumn = currentColumn + 1
                                        End If
                                End Select
                        Else
                            Select Case m_insertMode
                                ' Insert column before match.
                                Case 0
                                    .Item(currentColumn).Insert
                                    foundColumn = foundColumn + 1
                                ' Insert column after match.
                                Case 1
                                    currentColumn = currentColumn + 1
                                    .Item(currentColumn).Insert
                            End Select
                        End If
                    End With
                    ' Insert new category.
                    Do
                        currentRow = i + DataCategoryStartingRow
                        If GetCellValue(dataWS, currentRow, currentColumn) <> insertRows(i) Then
                            SetCellValue dataWS, currentRow, currentColumn, insertRows(i)
                            changed = True
                        End If
                        i = i + 1
                    Loop Until currentRow = DataCategoryStoppingRow
                End If
                ' List categories for file
                i = -m_listBreadth
                categoriesFile.Add dataWBFileName
                Do While i <= m_listBreadth
                    joinedCategoryName = vbNullString
                    currentColumn = foundColumn + i
                    currentRow = DataCategoryStartingRow
                    Do While currentRow <= DataCategoryStoppingRow
                        joinedCategoryName = joinedCategoryName & GetCellValue(dataWS, currentRow, currentColumn) & RowSep
                        currentRow = currentRow + 1
                    Loop
                    joinedCategoryName = Space$(4) & RemoveLast(joinedCategoryName, Len(RowSep))
                    categoriesFile.Add joinedCategoryName
                    i = i + 1
                Loop
                ' Exit category loop after match, only first match is relevant.
                Exit Do
            End If
            currentColumn = currentColumn + 1
        Loop
        ' Finish processing.
        If changed Then
            dataWB.Save
            changed = False
        End If
        dataWB.Close
Skip:
        dataWBFileName = Dir
        DoEvents
    Loop
    
    ' Save results in txt.
    Dim fileNumber As Long
    fileNumber = FreeFile()
    Dim entry As Variant
    Open ListFilesPath & m_listFileNameCategoryAdditions For Output As fileNumber
    For Each entry In categoriesFile
        Print #fileNumber, entry
    Next
    Close fileNumber
End Sub
