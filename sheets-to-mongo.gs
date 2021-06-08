// Create an object which contains keys for each column in the spreadsheet
var columns = { // 0 indexed
  name: 0,
  pendingresidents: 1,
  negativeresidents: 2,
  positiveresidents: 3,
  admittedwithcovid: 4,
  pendingstaff: 5,
  negativestaff: 6,
  positivestaff: 7,
  open: 8,
  dateUpdated: 9,
  locationId: 10
}

/****
 * This function runs automatically and adds a menu item to Google Sheets
 ****/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheetByName("Locations"));
  var entries = [{
    name : "Push COVID-19 Stats to Database",
    functionName : "exportLocationsToMongoDB"
  }];
  sheet.addMenu("Export COVID-19 Stats", entries);
};

/****
 * Export the events from the sheet to a MongoDB Database via Stitch
 ****/
function exportLocationsToMongoDB() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Locations");
  var headerRows = 1;  // Number of rows of header info (to skip)
  var range = sheet.getDataRange(); // determine the range of populated data
  var numRows = range.getNumRows(); // get the number of rows in the range
  var data = range.getValues(); // get the actual data in an array data[row][column]
  
  for (var i=headerRows; i<numRows; i++) {
    // Make a POST request with form data.
    var formData = {
      'name': data[i][columns.name],
      'pendingresidents': data[i][columns.pendingresidents],
      'negativeresidents': data[i][columns.negativeresidents],
      'positiveresidents': data[i][columns.positiveresidents],
      'pendingstaff': data[i][columns.pendingstaff],
      'negativestaff': data[i][columns.negativestaff],
      'positivestaff': data[i][columns.positivestaff],
      'admittedwithcovid': data[i][columns.admittedwithcovid],
      'open': data[i][columns.open],
      'dateUpdated': Utilities.formatDate(new Date(data[i][columns.dateUpdated]), "GMT", "MMM dd yyyy"),
      'locationId': data[i][columns.locationId]
    };
    var options = {
      'method' : 'post',
      'payload' : formData
    };
    var insertID = UrlFetchApp.fetch('[REDACTED]', options);
    //eventIdCell.setValue(""); // Insert the new event ID
  }
  
  for (var i=headerRows; i<numRows; i++) {
    var name = data[i][columns.name];
    var pendingresidents = data[i][columns.pendingresidents];
    var negativeresidents = data[i][columns.negativeresidents];     
    var positiveresidents = data[i][columns.positiveresidents];
    var pendingstaff = data[i][columns.pendingstaff];
    var negativestaff = data[i][columns.negativestaff];
    var positivestaff = data[i][columns.positivestaff];
    var admittedwithcovid = data[i][columns.admittedwithcovid];
    var open = data[i][columns.open];
    var dateUpdated = Utilities.formatDate(new Date(data[i][columns.dateUpdated]), "GMT", "MMM dd, yyyy");
    var locationId = data[i][columns.locationId];
    // Make a POST request with form data.
    var formData = {
      'name': data[i][columns.name],
      'pendingresidents': data[i][columns.pendingresidents],
      'negativeresidents': data[i][columns.negativeresidents],
      'positiveresidents': data[i][columns.positiveresidents],
      'pendingstaff': data[i][columns.pendingstaff],
      'negativestaff': data[i][columns.negativestaff],
      'positivestaff': data[i][columns.positivestaff],
      'admittedwithcovid': data[i][columns.admittedwithcovid],
      'open': data[i][columns.open],
      'dateUpdated': dateUpdated,
      'locationId': data[i][columns.locationId]
    };
    var options = {
      'method' : 'post',
      'payload' : formData
    };
    var insertID = UrlFetchApp.fetch('[REDACTED]', options);
    //eventIdCell.setValue(insertID); // Insert the new event ID
  }
}
