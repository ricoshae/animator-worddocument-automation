**Overview**

This code will automate the editing of word documents using AppleScript in Automator.
The script will request a folder of word documents then automatically open, edit and close each document based on simple find and replace commands in the script.

**To run the script**

Open Automator on a Mac.
Search for "Run AppleScript" and add to automator
Replace the example code with the script in animator-worddocument-automation.txt
Run the script
- Choose a folder
- script will open, edit and close each word document in the folder

To replacce content in the main docoument area use:
execute find (find object of docRange) find text "**jack**" replace with "**david**" replace replace all

To replace text in the header use:
- execute find (find object of headerRef) find text "**Songwriters book**" replace with "**Songwriting introduction**" replace replace all

To replace text in the footer use:
- execute find (find object of footerRef) find text "**2021**" replace with "**2025**" replace replace all

Remove the follwing line if you do not need to password protect the documents.
- protect active document protection type allow only form fields password "1234"




