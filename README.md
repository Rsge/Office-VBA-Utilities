# Office VBA Utilities
A collection of different VBA modules for different applications in MS Office.\
It's broken up by Office application, then Office file intention.\
A speed test section is also included to test different approaches' processing speeds.

All documents contain annotations used by [Rubberduck](https://rubberduckvba.com), an addon to the VBA editor adding lots of useful features, therefore "bringing it into the 21st century".\
I'd also recommmend making a few adjustments to the VBA editor's settings as seen in [this Stackoverflow comment](https://stackoverflow.com/a/667225/17239990).

### Excel
* `BasicUtilities` contains modules for automated clearing of specific part with prior confirmation and un-/protecting the document via bindable macro.
* `Finders/DifferingEntriesFinder.bas` lists all entries in a selected sheet's first column that aren't present in the first sheet's first column. Also copys results to clipboard.
* `Finders/UsedColumnsFinder.bas` lists all columns' headers in a text file, whose content isn't the same in every cell.
* `Fixes/ZipCodesFix.bas` fixes the US state abbreviation letters being put in front of the zip code instead of after the city in a table with column names.
* `Imports` contains modules for importing boolean or string values as text info for entries in an Excel sheet from a text or excel file.
* `InventoryUpdating` contains modules for importing data from multiple CSV files with relevant data on the last line into a table in excel, creating a backup of the current worksheet before making changes.
* `ProductionPlanning` contains modules for different calculations related to and needed for planning production, e.g. finding holidays and calculating production capacities.

### Outlook
* `BulkEditContacts.bas` allows for bulk-editing of contacts.
* `JunkMailBlackWhiteList.bas` force-whitelists mails under specified conditions and blacklists all addresses in a provided (linebreak-separated) text file (because "normal" rules are limited in max number of mails).
* `MoveFromFolderToFolder.bas` moves all mails from one folder to another reliably one-by-one.
* `MoveToCorrectSentFolder.cls` moves each sent email to the correct mailbox' Sent-elements-folder (Because per default they go into the main account's).

### Speed tests
* `AssignEmptyStringTest.bas` compares different methods to create an empty string.
* `DoesStringStartWithTest.bas` compares different methods to determine if a string starts with another string.
* `DoesStringEndWithTest.bas` compares different methods to determine if a string ends with another string.
* `IsStringEmptyTest.bas` compares different methods to determine if a string is empty.
* `IsStringInStringTest.bas` compares different methods to determine if a string contains another string.

### Word
* For `Continual numbering` all three files need to be used to achieve a counter counting the amount of pages printed on the page with the ability to set on print.
