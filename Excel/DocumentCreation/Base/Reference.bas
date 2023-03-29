Attribute VB_Name = "Reference"
Attribute VB_Description = "General constants."
'@Folder("DocumentCreation.Base")
'@ModuleDescription("General constants.")
Option Explicit

' String constants
'@VariableDescription("General separator for listed stuff.")
Public Const ListSep As String = ", "
Attribute ListSep.VB_VarDescription = "General separator for listed stuff."
'@VariableDescription("Separator of categories in line in document file.")
Public Const CategorySep As String = " / "
Attribute CategorySep.VB_VarDescription = "Separator of categories in line in document file."
'@VariableDescription("Separator of 'category ID' in line in document file.")
Public Const IDSep As String = " "
Attribute IDSep.VB_VarDescription = "Separator of 'category ID' in line in document file."
'@VariableDescription("Separator for lines in category names in data file.")
Public Const RowSep As String = " | "
Attribute RowSep.VB_VarDescription = "Separator for lines in category names in data file."
'@VariableDescription("File extension pattern of workbooks for data import.")
Public Const FileExt As String = ".xls?"
Attribute FileExt.VB_VarDescription = "File extension pattern of workbooks for data import."
'@VariableDescription("Path to workbooks for data import.")
Public Const DataWBsPath As String = "C:\Example\Data\"
Attribute DataWBsPath.VB_VarDescription = "Path to workbooks for data import."
'@VariableDescription("General pattern of workbooks for data import.")
Public Const AllDataWBFilesPattern As String = "*" & FileExt
Attribute AllDataWBFilesPattern.VB_VarDescription = "General pattern of workbooks for data import."
'@VariableDescription("Base path to files for listing categories.")
Public Const ListFilesPath As String = "C:\Example\Lists\"
Attribute ListFilesPath.VB_VarDescription = "Base path to files for listing categories."

' Localized strings
'@VariableDescription("Decimal symbol in native language document and in data files.")
Public Const DecimalSymbolNative As String = "."
Attribute DecimalSymbolNative.VB_VarDescription = "Decimal symbol in native language document and in data files."
'@VariableDescription("Decimal symbol in translated language document.")
Public Const DecimalSymbolTransl As String = ","
Attribute DecimalSymbolTransl.VB_VarDescription = "Decimal symbol in translated language document."
'@VariableDescription("Title of error messages.")
Public Const ErrorTitle As String = "Error!"
Attribute ErrorTitle.VB_VarDescription = "Title of error messages."
'@VariableDescription("Title of warning messages.")
Public Const WarningTitle As String = "Warning!"
Attribute WarningTitle.VB_VarDescription = "Title of warning messages."
'@VariableDescription("Ignore files with one of these strings in name. First entry also marks invalid sheets.")
Public Const IgnoreList As String = "old" & ListSep & "_"
Attribute IgnoreList.VB_VarDescription = "Ignore files with one of these strings in name. First entry also marks invalid sheets."
'@VariableDescription("Name of sheet with mapping of categories.")
Public Const MappingSheetName As String = "Assignments"
Attribute MappingSheetName.VB_VarDescription = "Name of sheet with mapping of categories."
'@VariableDescription("Name of table with mapping of document categories to data categories.")
Public Const DocToDataTableName As String = "DocToData"
Attribute DocToDataTableName.VB_VarDescription = "Name of table with mapping of document categories to data categories."
'@VariableDescription("Name of table with mapping of data categories to document categories.")
Public Const DataToDocTableName As String = "DatToDoc"
Attribute DataToDocTableName.VB_VarDescription = "Name of table with mapping of data categories to document categories."
'@VariableDescription("Name of table with entries in document categories region to skip over.")
Public Const DocSkipTableName As String = "DocSkip"
Attribute DocSkipTableName.VB_VarDescription = "Name of table with entries in document categories region to skip over."
'@VariableDescription("Error message for missing data header.")
Public Const MissingHeaderError As String = "No Header found." & vbNewLine & _
                                            "Please check data table manually."
Attribute MissingHeaderError.VB_VarDescription = "Error message for missing data header."
'@VariableDescription("Title of question for entry ID.")
Public Const EntryIDQuestionTitle As String = "Entry ID"
Attribute EntryIDQuestionTitle.VB_VarDescription = "Title of question for entry ID."
'@VariableDescription("Question for entry ID to import data from.")
Public Const EntryIDQuestion As String = "What is the entry ID?"
Attribute EntryIDQuestion.VB_VarDescription = "Question for entry ID to import data from."
'@VariableDescription("Error message for file with given ID not having been found at defined path. {}: 1 = File name, 2 = Path.")
Public Const FileNotFoundError As String = "The file ""{}"" was not found in:" & vbNewLine & "{}"
Attribute FileNotFoundError.VB_VarDescription = "Error message for file with given ID not having been found at defined path. {}: 1 = File name, 2 = Path."
'@VariableDescription("Error message for entry with given ID not having been found in open file. {}: 1 = Entry ID, 2 = File name.")
Public Const EntryNotFoundError As String = "The entry ""{}"" was not found in {}."
Attribute EntryNotFoundError.VB_VarDescription = "Error message for entry with given ID not having been found in open file. {}: 1 = Entry ID, 2 = File name."
'@VariableDescription("Error message for entry with given ID not having been found in open file. {}: 1 = Entry ID, 2 = File name.")
Public Const CategoryNotFoundWarning As String = "For the following categories, no value was found:" & vbNewLine
Attribute CategoryNotFoundWarning.VB_VarDescription = "Error message for entry with given ID not having been found in open file. {}: 1 = Entry ID, 2 = File name."
'@VariableDescription("Error message for duplicate entries in category definitions.")
Public Const DuplicateCategoryIDError As String = "For the entry ""{}"" the definition is added at multiple times."
Attribute DuplicateCategoryIDError.VB_VarDescription = "Error message for duplicate entries in category definitions."
'@VariableDescription("Error message for duplicate entries in category definitions.")
Public Const DecimalEntryError As String = "The entry ""{}"" for the decimal places column is wrong." & vbNewLine & _
                                           "Please only use whole numbers."
Attribute DecimalEntryError.VB_VarDescription = "Error message for duplicate entries in category definitions."

' - Rows & Columns -
' Mapping
'@VariableDescription("Index in DocToData table of column with category name in translated document.")
Public Const DocToDataTranslColum As Long = 1
Attribute DocToDataTranslColum.VB_VarDescription = "Index in DocToData table of column with category name in translated document."
'@VariableDescription("Index in DocToData table of column with category name in native document.")
Public Const DocToDataNativeColum As Long = 2
Attribute DocToDataNativeColum.VB_VarDescription = "Index in DocToData table of column with category name in native document."
'@VariableDescription("Index in DocToData table of column with category name in data.")
Public Const DocToDataDataColumn As Long = 3
Attribute DocToDataDataColumn.VB_VarDescription = "Index in DocToData table of column with category name in data."
'@VariableDescription("Index in DocToData table of column with decimal place info for data values.")
Public Const DocToDataDecimalColumn As Long = 4
Attribute DocToDataDecimalColumn.VB_VarDescription = "Index in DocToData table of column with decimal place info for data values."
'@VariableDescription("Index in DataToDoc table of column with entry string in data.")
Public Const DataToDocDataColumn As Long = 1
Attribute DataToDocDataColumn.VB_VarDescription = "Index in DataToDoc table of column with entry string in data."
'@VariableDescription("Index in DataToDoc table of column with entry string in translated document.")
Public Const DataToDocTranslColum As Long = 2
Attribute DataToDocTranslColum.VB_VarDescription = "Index in DataToDoc table of column with entry string in translated document."
'@VariableDescription("Index in DataToDoc table of column with entry string in native document.")
Public Const DataToDocNativeColum As Long = 3
Attribute DataToDocNativeColum.VB_VarDescription = "Index in DataToDoc table of column with entry string in native document."
' Doc
'@VariableDescription("Index of row with file ID in document.")
Public Const DocFileIDRow As Long = 1
Attribute DocFileIDRow.VB_VarDescription = "Index of row with file ID in document."
'@VariableDescription("Index of row with entry ID in document.")
Public Const DocEntryIDRow As Long = 2
Attribute DocEntryIDRow.VB_VarDescription = "Index of row with entry ID in document."
'@VariableDescription("Index of row with additional info in document.")
Public Const DocInfoRow As Long = 5
Attribute DocInfoRow.VB_VarDescription = "Index of row with additional info in document."
'@VariableDescription("Index of first row with categories and data in document.")
Public Const DocDataStartingRow As Long = 6
Attribute DocDataStartingRow.VB_VarDescription = "Index of first row with categories and data in document."
'@VariableDescription("Index of column with categories in document.")
Public Const DocCategoryColumn As Long = 1
Attribute DocCategoryColumn.VB_VarDescription = "Index of column with categories in document."
'@VariableDescription("Index of column with additional infos in document.")
Public Const DocInfosColumn As Long = 2
Attribute DocInfosColumn.VB_VarDescription = "Index of column with additional infos in document."
'@VariableDescription("Index of column with data in document.")
Public Const DocDataColumn As Long = 5
Attribute DocDataColumn.VB_VarDescription = "Index of column with data in document."
' Data
'@VariableDescription("Index of row with header in import data sheet.")
Public Const DataHeaderRow As Long = 1
Attribute DataHeaderRow.VB_VarDescription = "Index of row with header in import data sheet."
'@VariableDescription("Index of column with header in import data sheet.")
Public Const DataHeaderColumn As Long = 1
Attribute DataHeaderColumn.VB_VarDescription = "Index of column with header in import data sheet."
'@VariableDescription("Index of first row with category in import data sheet.")
Public Const DataCategoryStartingRow As Long = 3
Attribute DataCategoryStartingRow.VB_VarDescription = "Index of first row with category in import data sheet."
'@VariableDescription("Index of last row with category in import data sheet.")
Public Const DataCategoryStoppingRow As Long = 5
Attribute DataCategoryStoppingRow.VB_VarDescription = "Index of last row with category in import data sheet."
'@VariableDescription("Index of first column with data in import data sheet.")
Public Const DataStartingColumn As Long = 1
Attribute DataStartingColumn.VB_VarDescription = "Index of first column with data in import data sheet."

' Other integer constants
'@VariableDescription("How many following category ID parts to maximum try in combination.")
Public Const MaxCategoryIDCombinations As Long = 2
Attribute MaxCategoryIDCombinations.VB_VarDescription = "How many following category ID parts to maximum try in combination."
'@VariableDescription("Max length of each category value part to not remove spaces from the combined value.")
Public Const ReplaceSpaceThreshold As Long = 3
Attribute ReplaceSpaceThreshold.VB_VarDescription = "Max length of each category value part to not remove spaces from the combined value."
'@VariableDescription("Amount of rows to show when changing view.")
Public Const ShowRowsCount As Long = 5
Attribute ShowRowsCount.VB_VarDescription = "Amount of rows to show when changing view."
