// This script processes emails with attachments (per the QUERY value below),
// saves the attachments to Drive (per the DRIVE_FOLDER_ID value below), then
// recreates the emails without the attachments. Make sure to set those values
// to your liking before starting.
//
// See README.md for more info.

// The folder ID to fill with email attachments. Grab it from the URL like:
// https://drive.google.com/drive/u/1/folders/<DRIVE_FOLDER_ID>
const DRIVE_FOLDER_ID = "1lroGpgd78UCe3_yYZGF-dJt4CfD8l9Gc";
// A search query in Gmail.
const QUERY = "larger:5M before:2024";

function processMessages() {
  const msgList = Gmail.Users.Messages.list("me", { q: QUERY });
  let count = 0;
  Logger.log(`Estimated # messages to process: ${msgList.resultSizeEstimate}`);
  msgList.messages.forEach(function (m) {
    const email = GmailApp.getMessageById(m.id);
    const date = Utilities.formatDate(email.getDate(), Session.getScriptTimeZone()/*"GMT"*/, "yyyy-MM-dd");
    Logger.log(`Processing email #${++count} date: ${date} subject:"${email.getSubject()}"`);
    saveAttachmentsToDrive(email, DRIVE_FOLDER_ID);
    duplicateWithoutAttachments(m, email);
    moveToTrash(m.id);
    // throw new Error("debug stopping");
  });
  Logger.log("done");
}

function saveAttachmentsToDrive(email, driveFolderId) {
  const folder = DriveApp.getFolderById(driveFolderId);

  const date = Utilities.formatDate(email.getDate(), Session.getScriptTimeZone()/*"GMT"*/, "yyyy-MM-dd_HHmmss");
  const attachments = email.getAttachments();
  // Limit subject length in case of long subjects.
  const subject = email.getSubject().substring(0, 50);
  attachments.forEach(function (a) {
    folder.createFile(a.copyBlob()).setName(date + ' ' + subject + ' - ' + a.getName());
  });
}

// Find the end boundary of html or plain-text email
// Boundary looks something like: --=_715b.68 (must start with -- and be a whole line).
const reBoundaryPart = "([\r\n])(--[\\w=.-]*)[\r\n]*";
// Some emails have a Content-Transfer-Encoding declaration between the boundary and Content-Type.
const reContentTransferEncodingPart = "(Content-Transfer-Encoding: (7bit|quoted-printable|base64)[\r\n; ]*)?";
// Attachment sections may also have a Content-Disposition.
const reContentDispositionPart = "(Content-Disposition: inline[\r\n; ]*)?(filename=[^\r\n;]*)?"
// Supposed to end in semicolon but sometimes ends in a newline.
const reContentTypeHtmlPart = "Content-Type: text/html[;\r\n]";
const reContentTypePlainPart = "Content-Type: text/plain[;\r\n]";
const reContentTypeAttachmentPart = "Content-Type: (image/jpeg|application/pdf)[;\r\n]";
// Ignore case due to some emails using wrong case in tags.
const reHtml = new RegExp(reBoundaryPart + reContentTransferEncodingPart + reContentTypeHtmlPart, "i");
const rePlain = new RegExp(reBoundaryPart + reContentTransferEncodingPart + reContentTypePlainPart, "i");
const reAttachment = new RegExp(reBoundaryPart + reContentTransferEncodingPart + reContentDispositionPart + reContentTypeAttachmentPart, "i");

function duplicateWithoutAttachments(message, email) {
  // Get the `raw` email
  let raw = email.getRawContent();
  const threadId = message.getThreadId();
  const labelIds = message.getLabelIds();
  
  // Prefer HTML content if found.
  const reHtmlResult = reHtml.exec(raw);
  const reResult = reHtmlResult || rePlain.exec(raw);

  if (!reHtmlResult) {
    Logger.log("WARNING: non-html email! Seems more buggy. Check it worked.");
  }

  const DEBUG = false;

  if (!reResult || DEBUG) {
    for (var i = 0; i < raw.length; i += 2000) {
      const subs = raw.substring(i, i + 2000);
      if (subs.indexOf("text/plain") >= 0 || subs.indexOf("text/html") >= 0) {
        Logger.log("DEBUG: email chunk %d: %s", i, subs);
      }
    }
    if (!reResult) {
      throw new Error("ERROR: neither regex found anything!");
    }
  }

  // Find the index of the end-of-message boundary.
  // This will be equivalent to the boundary part we found, but later in the email,
  // and should be followed by "--".
  const start = reResult.index + reResult[0].length;
  let boundary = raw.indexOf(reResult[2] + "--", start);
  if (boundary < 0) {
    boundary = raw.indexOf(reResult[2], start);
    if (boundary < 0) {
      throw new Error("ERROR: boundary not found");
    }
  }

  // Get the part of the email up to the boundary.
  raw = raw.substring(0, boundary);

  if (raw.length > 1 * 1024 * 1024) {
    Logger.log(`WARNING: Email after attachment removal still >1MB. Size: ${raw.length}. Trying harder. Might need to run again to reprocess this email. Check it worked! <------------`);
    // Apple mail in particular seems to like reusing the same boundary ID for text and attachments.
    // Try finding and removing attachment sections within the boundary.
    let reAttachmentResult;
    while ((reAttachmentResult = reAttachment.exec(raw)) !== null) {
      if (DEBUG){
        Logger.log(`Found ${reAttachmentResult}, boundary ${reAttachmentResult[2]}.`);
      }
      // Find the next boundary.
      const attStart = reAttachmentResult.index + reAttachmentResult[0].length;
      const nextBoundary = raw.indexOf(reAttachmentResult[2], attStart);
      if (nextBoundary == -1) {
        Logger.log(`WARNING: couldn't find an end boundary for ${reAttachmentResult}`);
        break;
      }
      // Strip this section by keeping only the bits before and after the boundary string.
      const before = raw.substring(0, reAttachmentResult.index + reAttachmentResult[1].length);
      const after = raw.substring(nextBoundary, raw.length);
      if (DEBUG) {
        Logger.log(`Stripping out attachment from ${reAttachmentResult.index} to ${nextBoundary}`)
      }
      raw = before + after;
    }
    Logger.log(`Email size after more invasive method: ${raw.length}`);
  }

  let addendum = (reHtmlResult ? "<br><br>" : '\r\n\r\n') + "[Attachments stripped and downloaded]";
  if (raw.indexOf(addendum) > 0) {
    // Must have already processed this email? Don't add another addendum.
    addendum = "";
  }

  // Encode the attachment-free RFC 2822 formatted email string
  const base64_encoded_email = Utilities.base64EncodeWebSafe(raw + addendum);
  // Set the base64Encoded string to the `raw` required property
  const resource = { "raw": base64_encoded_email, "threadId": threadId, "labelIds": labelIds };

  // Re-insert the email with the original date/time 
  Gmail.Users.Messages.insert(resource, 'me', null, { 'internalDateSource': 'dateHeader' });
}

function moveToTrash(emailId) {
  Gmail.Users.Messages.trash('me', emailId);
}
