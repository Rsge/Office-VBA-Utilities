# Office VBA Utilities
A collection of different VBA modules for different applications in MS Office.\
It's broken up by Office application, then Office file intention.\
A speed test section is also included to test different approaches' processing speeds.

All documents contain annotations used by [Rubberduck](https://rubberduckvba.com), an addon to the VBA editor adding lots of useful features, therefore "bringing it into the 21st century".\
I'd also recommmend making a few adjustments to the VBA editor's settings as seen in [this Stackoverflow comment](https://stackoverflow.com/a/667225/17239990).

### Excel
* `BasicUtilities` contains modules for automated clearing of specific part with prior confirmation and un-/protecting the document via bindable macro.
* `Finders/DifferingEntriesFinder.bas` lists all entries in the active sheet's first column that aren't present in the first sheet's first column. Also copys results to clipboard.
* `Finders/EntryEqualizer.bas` deletes all entries in the active sheet's chosen column that aren't present in the first sheet's first column.
* `Finders/UsedColumnsFinder.bas` lists all columns' headers in a text file, whose content isn't the same in every cell.
* `Fixes/ZipCodesFix.bas` fixes the US state abbreviation letters being put in front of the zip code instead of after the city in a table with column names.
* `Imports` contains modules for importing boolean or string values as text info for entries in an Excel sheet from a text or excel file.
* `InventoryUpdating` contains modules for importing data from multiple CSV files with relevant data on the last line into a table in excel, creating a backup of the current worksheet before making changes.
* `ProductionPlanning` contains modules for different calculations related to and needed for planning production, e.g. finding holidays and calculating production capacities.

### Outlook
Modules with the `.cls` extension have to go into the `ThisOutlookSession` module or have to be called from there to work.
* `BulkEditContacts.bas` allows for bulk-editing of contacts.
* `DeleteSentFromSender.cls` deletes all mails of a specific sender address from the Sent folder permanently on Outlook startup.
* `JunkMailBlackWhiteList.bas` force-whitelists mails under specified conditions and blacklists all addresses in a provided (linebreak-separated) text file (because "normal" rules are limited in max number of mails).
* `MoveFromFolderToFolder.bas` moves all mails from one folder to another reliably one by one.
* `MoveToCorrectSentFolder.cls` moves each sent email to the correct mailbox' Sent Elements folder. (Because per default they go into the main account's.)\
There is also a [registry fix](https://github.com/Rsge/Windows-Error-Fixing-Scripts/blob/main/Set%20Outlook%20delegate%20sent%20items%20folder.reg) for this.

### Speed tests
* `AssignEmptyStringTest.bas` compares different methods to create an empty string.
* `AssignSpaceTest.bas` compares different methods to create a string of a space.
* `CompareNumberTest.bas` compares different methods to compare a number to another.
* `DoesStringStartWithTest.bas` compares different methods to determine if a string starts with another string.
* `DoesStringEndWithTest.bas` compares different methods to determine if a string ends with another string.
* `IsStringEmptyTest.bas` compares different methods to determine if a string is empty.
* `IsStringEqualToStringTest.bas` compares different methods to determine if a string equals another string.
* `IsStringInStringTest.bas` compares different methods to determine if a string contains another string.

### Word
* For `Sequentially numbered copies`, `ThisDocument.doccls` and either the `MultiPage` (MP) or `SinglePage` (SP) files need to be used to achieve a running number of printed pages on the page with the ability to set a starting value on print.
* `FileRegexReplace.bas` allows for bulk replacement of text in all files of a type in a folder and it's subfolders via multiple regexes.
