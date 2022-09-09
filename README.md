# Office VBA Utilities
A collection of different VBA modules for different applications in MS Office.\
It's broken up by Office application, then Office file intention.\
A speed test section is also included to test different approaches' processing speeds.\
\
All documents contain annotations used by [Rubberduck](https://rubberduckvba.com), an addon to the VBA editor adding lots of useful features, therefore "bringing it into the 21st century".\
I'd also recommmend making a few adjustments to the VBA editor's settings as seen in [this Stackoverflow comment](https://stackoverflow.com/a/667225/17239990).

### Excel
* `Cleanup & protection` contains methods for automated clearing of specific part with prior confirmation and un-/protecting the document via bindable macro.
* `Inventory updating` contains methods for importing data from multiple CSV files with relevant data on the last line into a table in excel, creating a backup of the current worksheet before making changes.
* `Production planning` contains methods for different calculations related to and needed for planning production, e.g. finding holidays and calculating production capacities.
* `Data importing` contains methods for importing boolean or string values as text info for entries in an Excel sheet from a text or excel file.
* `FindingUsedColumns.bas` lists all columns' headers in a text file, whose content isn't the same in every cell.

### Outlook
* `BulkEditContacts.bas` allows for bulk-editing of contacts.
* `JunkMailBlackWhiteList.bas` force-whitelists mails under specified conditions and blacklists all addresses in a provided (linebreak-separated) text file (because "normal" rules are limited in max number of mails).
* `MoveFromFolderToFolder.bas` moves all mails from one folder to another reliably one-by-one.
* `MoveToCorrectSentFolder.cls` moves each sent email to the correct mailbox' Sent-elements-folder (Because per default they go into the main account's).

### Speed tests
* `AssignEmptyStringTest.bas` compares using `x = ""` and `x = vbNullString` to create an empty string.
* `DoesStringStartWithTest.bas` compares using `InStr(x, y) = 1`, `InStrB(x, y) = 1`, `Left(x, Len(y)) = y`, `Left$(x, Len(y)) = y` and `x Like y*` to determine if a string starts with another string.
* `DoesStringEndWithTest.bas` compares using `InStr(StrReverse(x), z) = 1`, `InStrB(StrReverse(x), z) = 1`, `Right(x, Len(y)) = y`, `Right$(x, Len(y)) = y` and `x Like *y` to determine if a string ends with another string.
* `IsStringEmptyTest.bas` compares using `Len(x) = 0`, `LenB(x) = 0`, `x = ""` and `x = vbNullString` to determine if a string is empty.
* `IsStringInStringTest.bas` compares using `InStr(x, y) > 0`, `InStrB(x, y) > 0` and `x Like *y*` to determine if a string contains another string.

### Word
* For `Continual numbering` all three files need to be used to achieve a counter counting the amount of pages printed on the page with the ability to set on print.
