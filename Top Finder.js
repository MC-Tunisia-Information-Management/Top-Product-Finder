function findTopProducts() {
  var sheets = [
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
  ]; // List of sheet names to search
  var products = {}; // Object to store product counts

  // Loop through each sheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheets[i]);
    var lastRow = sheet.getLastRow();
    var productRange = sheet.getRange("J1:J" + lastRow);
    var productValues = productRange.getValues();

    // Loop through each product in the sheet
    for (var j = 0; j < productValues.length; j++) {
      var product = productValues[j][0];

      // Add product to object and increment count
      if (product !== "" && !products[product]) {
        products[product] = 1;
      } else if (product !== "") {
        products[product]++;
      }
    }
  }

  // Find products with highest counts
  var topProducts = [];
  var topCounts = [0, 0, 0];

  for (var product in products) {
    var count = products[product];

    if (count > topCounts[0]) {
      topCounts[2] = topCounts[1];
      topProducts[2] = topProducts[1];
      topCounts[1] = topCounts[0];
      topProducts[1] = topProducts[0];
      topCounts[0] = count;
      topProducts[0] = product;
    } else if (count > topCounts[1]) {
      topCounts[2] = topCounts[1];
      topProducts[2] = topProducts[1];
      topCounts[1] = count;
      topProducts[1] = product;
    } else if (count > topCounts[2]) {
      topCounts[2] = count;
      topProducts[2] = product;
    }
  }

  // Write output to cells J6, J7, and J8
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Welcome"); // Replace "Sheet1" with the name of your sheet
  sheet.getRange("H6").setValue(topProducts[0]);
  sheet.getRange("J6").setValue(topProducts[1]);
  sheet.getRange("L6").setValue(topProducts[2]);
}
