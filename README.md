# MailMerge

Basic Mail Merge from a Google Doc and Sheet

## Setup

1. Add this script to a Google Doc and run the `onOpen` function so that Google App Script prompts for approval.
2. Accept the permissions, then re-load the Google Doc. You should see a new menu item called "Email".
3. Write the template email, using {{variable1}}, {{variable2}}, etc. for the variables you want to replace.
4. Click "Create Spreadsheet" in the Email menu. This will create a new spreadsheet in your Google Drive.
5. Populate the fields of the Spreadsheet with the values you want to replace in the template email.
6. Open the Email menu of the Google Doc, and choose Create Drafts/Create Single Draft to confirm that the script is working.
7. Send the Drafts (or uncomment the "Send Email" option in onOpen() to send add a button in the Email menu to send immediately).
