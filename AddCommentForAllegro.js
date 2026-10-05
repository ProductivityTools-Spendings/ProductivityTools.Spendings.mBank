function fillAllegroPurchase(mbankTrixUrl, allegroTrixUrl) {
  // var trixUrl = "https://docs.google.com/spreadsheets/d/1XJAduyj-wL-kVE12Ib93htKbiEyTXuzYOG7j4BedrOA/edit?gid=0#gid=0"
  var sheet = getSheet(mbankTrixUrl, "mBankAccountExpenses")

  var dataRange = sheet.getDataRange();
  var rowsCount = dataRange.getNumRows();

  //var data = dataRange.getValues();
  for (i = 3500; i <= rowsCount; i++) {
    var description = dataRange.getCell(i, 10);
    var descriptionValue = description.getValue();
    console.log(descriptionValue)
    if (descriptionValue == "Allegro /Poznan" || descriptionValue == "ALLEGRO.PL" || descriptionValue == "ALLEGRO.PL &" || descriptionValue == "WWW.ALLEGRO.PL") {
      var whatCell = dataRange.getCell(i, 11);
      var note = whatCell.getNote()
      if (note == "") {
        var price = dataRange.getCell(i, 8).getValue();
        price = price * (-1)
        var allegroPurchasesWithGivenPrice = findByPrice(allegroTrixUrl, price)
        var formatedPurchases = formatNote(allegroPurchasesWithGivenPrice);
        whatCell.setNote(formatedPurchases)
      }
    }
  }
}

function formatNote(noteRaw) {
  var result = "";
  for (k = 0; k < noteRaw.length; k++) {
    result += noteRaw[k][0].toISOString().split('T')[0] + " " + noteRaw[k][1] + " " + noteRaw[k][2] + '\n';
  }
  console.log("Formateed:", result);
  return result;
}

// function findSomethging() {
//   var r = findByPrice("47.8")
//   console.log(r);
// }

function findByPrice(allegroTrixUrl, queryPrice) {
  // var allegroTrix = "https://docs.google.com/spreadsheets/d/1sBC7PWM7DkCA4smf11Gg59tWT5R7JRIRaWqqpkIYgw8/edit?gid=609545681#gid=609545681"
  var purchaseSheet = getSheet(allegroTrixUrl, "Purchases")
  var purchaseDataRange = purchaseSheet.getDataRange();
  var purchaseRowsCount = purchaseDataRange.getNumRows();
  var result = []

  for (j = 2; j <= purchaseRowsCount; j++) {
    var price = purchaseDataRange.getCell(j, 3);
    var priceValue = price.getValue();
    //console.log(priceValue);
    if (priceValue == queryPrice) {
      console.log("found");
      var itemName = purchaseDataRange.getCell(j, 4)
      var itemDate = purchaseDataRange.getCell(j, 1)
      var itemPrice = purchaseDataRange.getCell(j, 5)
      result.push([itemDate.getValue(), itemName.getValue(), itemPrice.getValue()])
      // var description = dataRange.getCell(i, 10).getValue();
      // if (categoryDictionary[description]) {
      //   category.setValue(categoryDictionary[description])
      // }
    }
  }
  return result;
}


function addNote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var targetCell = sheet.getRange("A1");
  var sourceCell = sheet.getRange("B1");

  var noteText = sourceCell.getValue();

  targetCell.setNote(noteText);

}