# Export MSG files from Outlook
## Purpose of this code
Exporting E-Mails from Outlook to .msg-files is easily done by marking and then drap-and-dropping the E-Mails from the Outlook window to the Windows file explorer.
However, this has several disadvantages for a large number of E-Mails:
- The process is very slow and no indication of the status is given.
- If the subject names of E-Mails have characters in them that are not valid in file names, the procedure fails.
- Sometimes, this will still fail unexpectedly.
The macro contained in this repo exports all E-Mails from a defined Outlook folder in a stable manner.

## Using this code
### Preperations
Before using this code, please ensure the following:
-  All items in the folder containing the E-Mails must be available offline. If they are not:
    1. Go to File -> Account Settings -> Account Settings -> double click on your account's name -> make all emails available offline
    2. Restart outlook
    3. Go to the folder -> Send/Receive -> Update folder
- If you are using a digital certificate, please log in first.
### Running the macro
Follow these steps to run the macro:
1. Open the VBA environment via `Alt+F11`
2. Import the macro: File -> Import file -> choose `ExportEmailToMsg.bas`
3. Open the immediate command window to view the status via `CTRL+G`
4. Run the macro via `F5`
5. First, you will be asked to select the Outlook folder to export E-Mails from. Then, you'll need to select a folder to save the msg-files to.


## Acknowledgements
Thanks to the authors and commentators on the following sites, who have inspired this code:
- https://stackoverflow.com/questions/57379087/save-outlook-email-to-my-internal-drive-as-msg-file
- https://www.extendoffice.com/documents/outlook/5034-outlook-save-multiple-emails-as-msg.html
