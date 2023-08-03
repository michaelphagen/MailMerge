# MailMerge

Basic Mail Merge from a Google Doc and Sheet

## Setup

1. Add this script to a Google Doc, then re-load the document. You should see a new menu item called "Email".
2. Write the template email, using {{variable1}}, {{variable2}}, etc. for the variables you want to replace.
3. Click "Create Spreadsheet" in the Email menu. This will create a new spreadsheet in your Google Drive.
4. Populate the fields of the Spreadsheet with the values you want to replace in the template email.
5. Open the Email menu of the Google Doc, and choose Create Drafts/Create Single Draft to confirm that the script is working.
6. Send the Drafts (or uncomment the "Send Email" option in onOpen() to send add a button in the Email menu to send immediately).
