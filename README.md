# Office VBA Utilities
A collection of different VBA scripts for different applications in MS Office.\
It's broken up by Office application, then Office file intention.\
A speed test section is also included to test different approaches' processing speeds.\
\
All documents contain annotations used by [Rubberduck](https://rubberduckvba.com), an addon to the VBA editor adding lots of useful features.\
I'd also recommmend making a few adjustments to the VBA editor's settings as seen in [this Stackoverflow comment](https://stackoverflow.com/a/667225/17239990).

### Excel
* `Attendence time table` has scripts for automated clearing of specific part with prior confirmation and un-/protecting the document via bindable macro.
* `Production planning` has lots of scripts for different calculations related to and needed for planning production, e.g. finding holidays and calculating production capacities.

### Outlook
* `BulkEditContacts.bas` allows for bulk-editing of contacts.
* `JunkMailBlackWhiteList.bas` force-whitelists mails under specified conditions and blacklists all addresses in a provided (linebreak-separated) text file (because "normal" rules are limited in max number of mails).
* `MoveFromFolderToFolder.bas` moves all mails from one folder to another reliably one-by-one.
* `MoveToCorrectSentFolder.cls` moves each sent email to the correct mailbox' Sent-elements-folder (Because per default they go into the main account's).

### Speed tests
* `AssignEmptyStringTest.bas` compares using `x = ""` and `x = vbNullString` to create an empty string.
* `IsStringEmptyTest.bas` compares using `Len(x) = 0`, `LenB(x) = 0`, `x = ""` and `x = vbNullString` to determine if a string is empty.

### Word
* For `Continual numbering` all three files need to be used to achieve a counter counting the amount of pages printed on the page with the ability to set on print.
