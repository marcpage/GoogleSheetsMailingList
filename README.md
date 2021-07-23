# GoogleSheetsMailingList
Attach this script to a google sheet and you'll be able to send emails to mailing lists in the sheet

## How to install

1. Open a Google Sheet and select Tools -> Script Editor
2. Replace the contents of the Code.gs file with the script above
3. Click the "Save" icon (ctl-s or command-s)
4. Select the onOpen function as the function to run and run it
5. You will be asked for permissions to Sheets and to Gmail, accept

## How to use

Set up a sheet whose name (the name of the tab at the bottom) is the name of an email draft you have.
On that sheet, add names and email addresses (names in first column, email addresses in second column) who you'd like to send the email to.
Notice a new "Actions" menu. Select Actions -> Send Emails.
Once the script is done, you will see "Sent" in the 3rd column.

The emails will be duplicates of the draft with the subject that matches the name of the tab.
The To: field will be the email address in column 2.
The first line of the email will be the name in column 1 followed by a comma.

The email sent will look just like the draft email (including CC and BCC fields, images in the message, and attachments).
