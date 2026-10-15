const MBANK_TRANSFERS_LABEL = ".Archive/mBankTransfers";
// Gmail batch APIs accept at most 100 threads per call
const MOVE_BATCH_SIZE = 100;
// Set to false if you want the threads to keep the label after moving them to the inbox
const REMOVE_LABEL_AFTER_MOVE = true;

/**
 * Moves every thread labelled ".Archive/mBankTransfers" back to the inbox,
 * so that exportMBankEmailsToSpendingsApi() picks them up again (it will re-label and re-archive them).
 *
 * Safe to run multiple times: already moved threads simply no longer carry the label.
 * If the run approaches the Apps Script execution time limit it stops early and logs how much is left -
 * just run it again.
 */
function moveMBankTransfersToInbox() {
  var label = GmailApp.getUserLabelByName(MBANK_TRANSFERS_LABEL);
  if (!label) {
    console.error("Label not found:", MBANK_TRANSFERS_LABEL);
    return;
  }

  var startedAt = Date.now();
  var maxRuntimeMs = 5 * 60 * 1000; // leave a margin below the 6 min limit
  var movedTotal = 0;

  while (true) {
    // Always fetch from offset 0: once a batch loses the label it disappears from this query.
    // With REMOVE_LABEL_AFTER_MOVE=false we instead have to page through with an offset.
    var offset = REMOVE_LABEL_AFTER_MOVE ? 0 : movedTotal;
    var threads = label.getThreads(offset, MOVE_BATCH_SIZE);
    if (threads.length === 0) {
      break;
    }

    GmailApp.moveThreadsToInbox(threads);
    if (REMOVE_LABEL_AFTER_MOVE) {
      label.removeFromThreads(threads);
    }

    movedTotal += threads.length;
    console.log("Moved to inbox so far:", movedTotal);

    if (Date.now() - startedAt > maxRuntimeMs) {
      console.warn("Stopping before execution time limit. Run the script again to move the remaining threads.");
      break;
    }
  }

  console.log("Done. Threads moved to inbox:", movedTotal);
}
