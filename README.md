# RemoveGmailAttachments
A Google Apps Script for removing attachments from emails in Gmail.

This script processes emails with attachments (per the QUERY value in the script),
saves the attachments to Drive (per the DRIVE_FOLDER_ID value in the script), then
recreates the emails without the attachments.

## Setup
To set up:
- Go to: https://script.google.com/ and make a new project.
- Put the RemoveGmailAttachments script in the project (can just paste it in).
- Click the + next to Services and add the Gmail service.
- Create a folder in Drive and paste the ID from the URL into DRIVE_FOLDER_ID into the script.
- Optional: change the QUERY value in the script if desired.
- Click Run

## Note
Email replies often include attachments from the original message,
leading to duplicate attachments being exported, so you should run a
deduplication check on the exported files.

## Warning
Running this will move the processed original emails to your Gmail trash bin.
I strongly suggest hand-checking some of the processed emails
(especially any warned about in the execution log) to see that it operated correctly.
It's fundamentally just a hacky regex trying to parse emails so it's bound to have
issues on some emails.
