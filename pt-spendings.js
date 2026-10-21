const SPENDINGS_WEBAPI_URL = "https://spendings-api.productivitytools.top/api/emails/raw";
const MBANK_SUBJECT = "mBank - powiadomienie e-mail";
const MBANK_ATTACHMENT_PREFIX = "Powiadomienie e-mail z";
const MBANK_ARCHIVE_LABEL = ".Archive/mBankTransfers";
// The Inbox is expected to hold at most ~100 notifications between runs, so one search page is enough.
const MAX_THREADS = 500;

/**
 * Sends mBank notification attachments from the Inbox to the Spendings WebApi, oldest first,
 * so that ids in the database follow the e-mail chronology. Each thread is archived (and labelled)
 * right after all its messages were accepted; on the first failed request the run stops, because
 * continuing would insert newer e-mails before an older one - fix the cause and run again.
 */
function exportMBankEmailsToSpendingsApi() {
  var messages = collectMBankMessagesFromInbox();
  if (messages.length === 0) {
    console.log("No mBank notifications in Inbox.");
    return;
  }
  messages.sort(function (a, b) {
    return a.date - b.date || a.messageId.localeCompare(b.messageId);
  });
  console.log("Found", messages.length, "messages. Sending oldest first...");

  var label = GmailApp.getUserLabelByName(MBANK_ARCHIVE_LABEL);
  var sent = 0;
  for (var i = 0; i < messages.length; i++) {
    var m = messages[i];
    var attachments = m.message.getAttachments().filter(function (a) {
      return a.getName().startsWith(MBANK_ATTACHMENT_PREFIX);
    });
    for (var a = 0; a < attachments.length; a++) {
      if (!sendAttachmentToSpendingsApi(m, attachments[a])) {
        console.warn("Stopped on a failed request. Sent", sent, "of", messages.length,
          "messages; the rest stays in Inbox - run again after fixing the cause.");
        return;
      }
    }
    sent++;
    // The thread is archived once its last (newest) message has been sent.
    if (m.lastInThread) {
      if (label) {
        m.thread.addLabel(label);
      }
      m.thread.moveToArchive();
    }
  }
  console.log("Done. Messages sent:", sent);
}

/** Returns every message of the mBank notification threads in the Inbox (one Gmail search). */
function collectMBankMessagesFromInbox() {
  var threads = GmailApp.search('in:inbox subject:"' + MBANK_SUBJECT + '"', 0, MAX_THREADS);
  if (threads.length === MAX_THREADS) {
    console.warn("Inbox has", MAX_THREADS, "or more mBank threads - only the first page is processed; run again.");
  }
  var result = [];
  threads.forEach(function (thread) {
    if (thread.getFirstMessageSubject() !== MBANK_SUBJECT) {
      return;
    }
    var msgs = thread.getMessages();
    msgs.forEach(function (msg, i) {
      result.push({
        thread: thread,
        threadId: thread.getId(),
        message: msg,
        messageId: msg.getId(),
        date: msg.getDate().getTime(),
        subject: msg.getSubject(),
        lastInThread: i === msgs.length - 1
      });
    });
  });
  return result;
}

/** POSTs one attachment to the WebApi. Returns true when the API accepted it (2xx). */
function sendAttachmentToSpendingsApi(m, attachment) {
  var payload = {
    source: "MBANK",
    threadId: m.threadId,
    messageId: m.messageId,
    subject: m.subject,
    attachmentName: attachment.getName(),
    emailDate: new Date(m.date).toISOString(),
    rawHtml: attachment.getDataAsString()
  };
  var response = UrlFetchApp.fetch(SPENDINGS_WEBAPI_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    console.log("Sent:", new Date(m.date).toISOString(), m.messageId, attachment.getName());
    return true;
  }
  console.error("Failed to send attachment. Status:", code, "Body:", response.getContentText(), "messageId:", m.messageId);
  return false;
}

/**
 * HTTP entry point so the export can be started from a link in the browser:
 *   GET https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
 *
 * Deploy as "Web app" with Execute as: Me and Who has access: Only myself - Google then lets only the
 * owner's signed-in browser session reach this function, so no extra secret is needed. Apps Script's 6-minute
 * execution limit applies; with ~100 e-mails per run that is far away.
 */
function doGet() {
  var startedAt = Date.now();
  try {
    exportMBankEmailsToSpendingsApi();
    return textResponse("OK: exportMBankEmailsToSpendingsApi finished in "
      + Math.round((Date.now() - startedAt) / 1000) + " s. See Executions in Apps Script for the log.");
  } catch (err) {
    console.error("doGet failed:", err);
    return textResponse("ERROR: " + err);
  }
}

function textResponse(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}
