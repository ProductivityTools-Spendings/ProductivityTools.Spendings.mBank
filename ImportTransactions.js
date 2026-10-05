function importTransactionsFromGmail(trixUrl) {
  var threads = GmailApp.getInboxThreads();
  for (var i = threads.length - 1; i >= 0; i--) {
    var thread = threads[i];
    var subject = thread.getFirstMessageSubject();
    if (subject == "mBank - powiadomienie e-mail") {
      processOneMBankMessage(trixUrl, thread);
    }
  }
}

function getTransferRow(threadOperationId, date, data) {
  var details = data[1]
  if (details.startsWith("mBank: Autoryzacja karty")) {
    var segments = details.split(" ");
    var srcCard = segments[3].replace(":", "").trim();
    segments = details.split(":")
    var name = segments[2].replace(". Kwota", "").trim();
    var amountRaw = segments[3].trim();
    var amount = -1 * Number(amountRaw.split(" ")[0].replace(",", "."));
    var currency = amountRaw.split(" ")[1].replace(".", "");
    var amountLeftRaw = segments[4].trim();
    var amountLeft = amountLeftRaw.split(" ")[0]
    var amountLeftCurrency = amountLeftRaw.split(" ")[1].replace(".", "");

    //var segments = details.split(":");
    return [threadOperationId, date, data[0], "Autoryzacja karty", srcCard, "", "", amount, currency, name, "", "", amountLeft, amountLeftCurrency, '', details]
  }
  if (details.startsWith("mBank: Przelew wych")) {
    var segments = details.split(" ");
    var srcAccount = segments[5];
    var dstAccount = segments[8];
    var amount = -1 * Number(segments[10].replace(",", "."));
    var currency = segments[11];
    var nameSegments = details.split("dla")[1].split("Dost.")[0].trim();
    if (details.includes("Dost.")) {
      var left = details.split("dla")[1].split("Dost.")[1].trim().split(" ");
      var leftValue = left[0];
      var leftCurrency = left[1].replace(".", "");
    }
    else
    {
      var leftValue="??"
      var leftCurrency="??"
    }
    return [threadOperationId, date, data[0], "Przelew wychodzacy", srcAccount, dstAccount, "", amount, currency, nameSegments, "", "", leftValue, leftCurrency, '', details]
  }
  if (details.startsWith("mBank: Przelew przych")) {
    var segments = details.split(" ");
    var srcAccount = segments[5];
    var dstAccount = segments[8];
    var amount = 1 * Number(segments[10].replace(",", "."));
    var currency = segments[11];
    var name = details.split(" od ")[1].split("Dost.")[0].trim();
    var temp = details.split(" o d");
    Logger.log(temp);
    temp = details.split(" od ")[1];
    Logger.log(temp);
    temp = details.split(" od ")[1].split("Dost.");
    Logger.log(temp);
    var left = details.split(" od ")[1].split("Dost.")[1].trim().split(" ");
    var leftValue = left[0];
    var leftCurrency = left[1].replace(".", "");;
    return [threadOperationId, date, data[0], "Przelew przychodzący", srcAccount, dstAccount, "", amount, currency, name, "", "", leftValue, leftCurrency, '', details]
  }

  if (details.startsWith("mBank: Obciazenie")) {
    var segments = details.split(" ");
    var srcAccount = segments[3]
    var amount = -1 * Number(segments[6].replace(",", "."));
    var amountCurrency = segments[7]
    var name = details.split('tytulem: ')[1].split(";")[0].trim();
    var amountLeftRaw = details.split('tytulem: ')[1].split(";")[1].trim().split(" ")
    var amountLeft = amountLeftRaw[1];
    var amountLeftCurrency = amountLeftRaw[2]
    return [threadOperationId, date, data[0], "Obciazenie", srcAccount, "", "", amount, amountCurrency, name, "", "", amountLeft, amountLeftCurrency, '', details]
  }

  if (details.startsWith("mBank: Uznanie na rach.")) {
    var segments = details.split(" ");
    var dstAccount = segments[4]
    var amount = Number(segments[7].replace(",", "."));
    var amountCurrency = segments[8]
    var name = details.split('tytulem: ')[1].split(";")[0].trim();
    var amountLeftRaw = details.split('tytulem: ')[1].split(";")[1].trim().split(" ")
    var amountLeft = amountLeftRaw[1];
    var amountLeftCurrency = amountLeftRaw[2]
    return [threadOperationId, date, data[0], "Uznanie", "", dstAccount, "", amount, amountCurrency, name, "", "", amountLeft, amountLeftCurrency, '', details]
  }
  if (details.startsWith("mBank: Odmowa autoryzacji") ||
    details.startsWith("mBank: Potwierdzenie poprawnego") ||
    details.split('mBank: Twoj przelew do') ||
    details.split('mBank: Niepoprawne logowanie ')) {
    return null;
  }


  return [date, data[0], "xx", details]
}


function copyDataToSpreadsheet(trixUrl, threadId, date, data) {
  var sheetAccountExpenses = getSheet(trixUrl, "mBankAccountExpenses")
  var sheetCardExpenses = getSheet(trixUrl, "mBankCardExpenses")
  for (var row = 0; row < data.length; row++) {
    if (data[row][0].startsWith("Czas") == false) {
      var threadOperationId = threadId + '-' + row;
      var transferRow = getTransferRow(threadOperationId, date, data[row]);
      if (transferRow != null) {

        if (transferRow[3] == "Autoryzacja karty") {
          sheetCardExpenses.appendRow(transferRow)
        }
        else {
          sheetAccountExpenses.appendRow(transferRow)
        }
      }
    }
  }
}

function getDateFromEmailAttachement(attachement) {
  const content = attachement.getDataAsString();
  const $ = Cheerio.load(content);
  var x = $("table > tbody > tr > td > h1").first();
  var date = x.text().replace("Powiadomienie e-mail", "").trim().replace(" -", "");
  return date;
}

function getDataFromEmailAttachement(trixUrl, threadId, attachement) {
  //console.log(attachement.getDataAsString())
  const content = attachement.getDataAsString();
  const $ = Cheerio.load(content);
  var data = [];
  $("table > tbody > tr > td > table > tbody > tr").each((index, element) => {
    var row = [];
    $(element).find("td").each((index, child) => {
      var rowContent = $(child).text().trim()
      row.push(rowContent);
    });
    if (row.length > 0) {
      data.push(row);
    }
  });

  var date = getDateFromEmailAttachement(attachement);
  copyDataToSpreadsheet(trixUrl, threadId, date, data);
  console.log(data);
}

function processOneMBankMessage(trixUrl, thread) {
  var msgs = thread.getMessages();
  var threadId = thread.getId();
  var mbank = false;
  var label = GmailApp.getUserLabelByName(".Archive/mBankTransfers");

  for (var j = 0; j < msgs.length; j++) {
    var attachments = msgs[j].getAttachments();
    for (var k = 0; k < attachments.length; k++) {
      var attachement = attachments[k]
      var attachementName = attachement.getName();
      if (attachementName.startsWith("Powiadomienie e-mail z")) {
        //console.log(attachementName)
        getDataFromEmailAttachement(trixUrl, threadId, attachement);
        mbank = true;
      }
      //SaveAttachement(attachments[k], targetDirectory)
    }
    if (mbank) {
      thread.addLabel(label);
      thread.moveToArchive();

    }
  }
}


