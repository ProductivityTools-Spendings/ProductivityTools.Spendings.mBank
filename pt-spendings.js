const SPENDINGS_WEBAPI_URL = "http://34.116.163.207:8086/api/emails/raw";

function exportMBankEmailsToSpendingsApi() {
  var threads = GmailApp.getInboxThreads();
  for (var i = threads.length - 1; i >= 0; i--) {
    var thread = threads[i];
    var subject = thread.getFirstMessageSubject();
    if (subject == "mBank - powiadomienie e-mail") {
      sendOneMBankThreadToSpendingsApi(thread);
    }
  }
}

function sendOneMBankThreadToSpendingsApi(thread) {
  var msgs = thread.getMessages();
  var threadId = thread.getId();
  var label = GmailApp.getUserLabelByName(".Archive/mBankTransfers");
  var allSentSuccessfully = false;

  for (var j = 0; j < msgs.length; j++) {
    var msg = msgs[j];
    var messageId = msg.getId();
    var emailDate = msg.getDate().toISOString();
    var subject = msg.getSubject();
    var attachments = msg.getAttachments();

    for (var k = 0; k < attachments.length; k++) {
      var attachment = attachments[k];
      var attachmentName = attachment.getName();

      if (attachmentName.startsWith("Powiadomienie e-mail z")) {
        var rawHtml = attachment.getDataAsString();

        var payload = {
          source: "MBANK",
          threadId: threadId,
          messageId: messageId,
          subject: subject,
          attachmentName: attachmentName,
          emailDate: emailDate,
          rawHtml: rawHtml
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
          console.log("Successfully sent attachment to WebApi:", messageId, attachmentName);
          allSentSuccessfully = true;
        } else {
          console.error(
            "Failed to send attachment to WebApi. Status:",
            responseCode,
            "Body:",
            response.getContentText()
          );
          return;
        }
      }
    }
  }

  if (allSentSuccessfully) {
    if (label) {
      thread.addLabel(label);
    }
    thread.moveToArchive();
  }
}
