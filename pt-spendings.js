const SPENDINGS_WEBAPI_URL = "http://34.116.163.207:8086/api/emails/raw";
const MBANK_SUBJECT = "mBank - powiadomienie e-mail";
const MBANK_ATTACHMENT_PREFIX = "Powiadomienie e-mail z";
const MBANK_ARCHIVE_LABEL = ".Archive/mBankTransfers";
const INBOX_PAGE_SIZE = 100;
// Apps Script kills the execution after 6 minutes; stop cleanly before that and continue on the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;

/**
 * Sends mBank notification attachments from the Inbox to the Spendings WebApi,
 * strictly oldest-first, so that ids in the database follow the e-mail chronology.
 *
 * - Pass 1 (cheap): lists Inbox threads (paged) and collects message dates only.
 * - Pass 2: in date order reads the attachment, POSTs it, and archives the thread
 *   as soon as all of its messages were accepted - progress is kept even if the run is cut short.
 * - Stops at the first failed request: continuing would insert newer e-mails before an older one.
 * - Stops before the execution time limit; simply run again to continue with the remaining e-mails.
 */
function exportMBankEmailsToSpendingsApi() {
  var startedAt = Date.now();
  var messages = collectMBankMessagesFromInbox();
  if (messages.length === 0) {
    console.log("No mBank notifications in Inbox.");
    return;
  }

  messages.sort(function (a, b) {
    return a.date - b.date || a.messageId.localeCompare(b.messageId);
  });

  var remainingPerThread = {};
  var threadsById = {};
  messages.forEach(function (m) {
    remainingPerThread[m.threadId] = (remainingPerThread[m.threadId] || 0) + 1;
    threadsById[m.threadId] = m.thread;
  });
  console.log("Found", messages.length, "messages in", Object.keys(threadsById).length,
    "threads (listing took", Math.round((Date.now() - startedAt) / 1000), "s). Sending oldest first...");

  var label = GmailApp.getUserLabelByName(MBANK_ARCHIVE_LABEL);
  var sentMessages = 0;
  var sentAttachments = 0;
  var archived = 0;
  var stoppedReason = null;

  for (var i = 0; i < messages.length; i++) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      stoppedReason = "time limit";
      break;
    }

    var m = messages[i];
    var attachments = m.message.getAttachments();
    var failed = false;
    for (var a = 0; a < attachments.length; a++) {
      var attachment = attachments[a];
      if (!attachment.getName().startsWith(MBANK_ATTACHMENT_PREFIX)) {
        continue;
      }
      if (!sendAttachmentToSpendingsApi(m, attachment)) {
        failed = true;
        break;
      }
      sentAttachments++;
    }
    if (failed) {
      stoppedReason = "failed request";
      break;
    }

    sentMessages++;
    // Archive the thread right away once every message in it has been sent.
    remainingPerThread[m.threadId]--;
    if (remainingPerThread[m.threadId] === 0) {
      var thread = threadsById[m.threadId];
      if (label) {
        thread.addLabel(label);
      }
      thread.moveToArchive();
      archived++;
    }
  }

  if (stoppedReason) {
    console.warn("Stopped early (" + stoppedReason + "). Sent", sentMessages, "of", messages.length,
      "messages; the rest stays in Inbox - run the script again to continue.");
  }
  console.log("Done. Messages sent:", sentMessages, "/", messages.length,
    "; attachments sent:", sentAttachments, "; threads archived:", archived);
}

/**
 * Walks the whole Inbox page by page and returns mBank messages with their dates.
 * Attachments are NOT read here (that is the slow part) - they are fetched right before sending.
 */
function collectMBankMessagesFromInbox() {
  var result = [];
  var start = 0;
  while (true) {
    var threads = GmailApp.getInboxThreads(start, INBOX_PAGE_SIZE);
    if (threads.length === 0) {
      break;
    }
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      if (thread.getFirstMessageSubject() !== MBANK_SUBJECT) {
        continue;
      }
      var msgs = thread.getMessages();
      for (var i = 0; i < msgs.length; i++) {
        result.push({
          thread: thread,
          threadId: thread.getId(),
          message: msgs[i],
          messageId: msgs[i].getId(),
          date: msgs[i].getDate().getTime(),
          subject: msgs[i].getSubject()
        });
      }
    }
    start += threads.length;
  }
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

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(SPENDINGS_WEBAPI_URL, options);
  var responseCode = response.getResponseCode();
  if (responseCode >= 200 && responseCode < 300) {
    console.log("Sent:", new Date(m.date).toISOString(), m.messageId, attachment.getName());
    return true;
  }
  console.error("Failed to send attachment. Status:", responseCode, "Body:", response.getContentText(),
    "messageId:", m.messageId);
  return false;
}


/**
 * HTTP entry point so the export can be started from a link in the browser:
 *   GET https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
 *
 * Deploy as "Web app" with Execute as: Me and Who has access: Only myself - Google then lets only the
 * owner's signed-in browser session reach this function, so no extra secret is needed. Same 6-minute limit
 * as a manual run: the export stops itself after ~5 minutes, open the URL again to continue.
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
