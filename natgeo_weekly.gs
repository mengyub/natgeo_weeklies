function RunAll() {
  // Run both repoorts
  RunWeeklySalesReport()
  RunWeeklyReconciliationReport()
}


// WEEKLY SALES
function RunWeeklySalesReport() {
  var fSource = DriveApp.getFolderById('13tDzrmLj1XKJ0W6d9U2YxyKSPDn9C6IW'); // id of folder where csv reports are saved
  var fi = fSource.getFilesByName('weekly_sales.csv'); // latest report file
  var ss = SpreadsheetApp.openById('1WBNk3VmmXmH4P0m_0tfdQTcknO7kdge87Kst95aBrFU'); // id of spreadsheet that holds the data to be updated with new report data

  if ( fi.hasNext() ) { // proceed if "weekly_sales.csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); // see below for CSVToArray function
    var dataSheet = ss.getSheetByName('NEWDATA') 
    // delete all existing rows
    var row_del = new Array();
    for (var i = row_del.length - 1; i>=0; i--) {
      dataSheet.deleteRow(row_del[i]); 
    }
    // var newsheet = ss.insertSheet('NEWDATA'); // create a 'NEWDATA' sheet to store imported data
    // loop through csv data array and insert (append) as rows into 'NEWDATA' sheet
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      dataSheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    // rename the report.csv file so it is not processed on next scheduled run
    file.setName("sales data imported on "+(new Date().toString())+".csv");
    SpreadsheetApp.flush();
    CopyWeeklySalesReport();
  }
};

function CopyWeeklySalesReport() {
  var ReportDestFolder = DriveApp.getFolderById('11jwq_eL5pOtspMN9WDUmxJJ3uUuKUex1'); 
  var ReportTemplate = DriveApp.getFileById('1WBNk3VmmXmH4P0m_0tfdQTcknO7kdge87Kst95aBrFU');
  ReportTemplate.makeCopy("NatGeo Weekly Sales Report - "+(new Date().toString()), ReportDestFolder);
};


// WEEKLY RECONCILIATION
function RunWeeklyReconciliationReport() {
  var fSource = DriveApp.getFolderById('13tDzrmLj1XKJ0W6d9U2YxyKSPDn9C6IW'); // id of folder where csv reports are saved
  var fi = fSource.getFilesByName('weekly_reconciliation.csv'); // latest report file
  var ss = SpreadsheetApp.openById('1_EIJDgkDtxuEqfQFqUMQpF3oMRf8ng4INwBmJAvdXJA'); // id of spreadsheet that holds the data to be updated with new report data

  if ( fi.hasNext() ) { // proceed if "weekly_reconciliation.csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); // see below for CSVToArray function
    var dataSheet = ss.getSheetByName('NEWDATA') 
    // delete all existing rows
    var row_del = new Array();
    for (var i = row_del.length - 1; i>=0; i--) {
      dataSheet.deleteRow(row_del[i]); 
    }
    // var newsheet = ss.insertSheet('NEWDATA'); // create a 'NEWDATA' sheet to store imported data
    // loop through csv data array and insert (append) as rows into 'NEWDATA' sheet
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      dataSheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    // rename the report.csv file so it is not processed on next scheduled run
    file.setName("reconciliation data imported on "+(new Date().toString())+".csv");
    SpreadsheetApp.flush();
    CopyWeeklyReconciliationReport();
  }
};

function CopyWeeklyReconciliationReport() {
  var ReportDestFolder = DriveApp.getFolderById('11jwq_eL5pOtspMN9WDUmxJJ3uUuKUex1'); 
  var ReportTemplate = DriveApp.getFileById('1_EIJDgkDtxuEqfQFqUMQpF3oMRf8ng4INwBmJAvdXJA');
  ReportTemplate.makeCopy("NatGeo Weekly Reconciliation Report - "+(new Date().toString()), ReportDestFolder);
};


// WEEKLY DEPARTURES
function RunWeeklyReconciliationReport() {
  var fSource = DriveApp.getFolderById('13tDzrmLj1XKJ0W6d9U2YxyKSPDn9C6IW'); // id of folder where csv reports are saved
  var fi = fSource.getFilesByName('weekly_reconciliation.csv'); // latest report file
  var ss = SpreadsheetApp.openById('1_EIJDgkDtxuEqfQFqUMQpF3oMRf8ng4INwBmJAvdXJA'); // id of spreadsheet that holds the data to be updated with new report data

  if ( fi.hasNext() ) { // proceed if "weekly_reconciliation.csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); // see below for CSVToArray function
    var dataSheet = ss.getSheetByName('NEWDATA') 
    // delete all existing rows
    var row_del = new Array();
    for (var i = row_del.length - 1; i>=0; i--) {
      dataSheet.deleteRow(row_del[i]); 
    }
    // var newsheet = ss.insertSheet('NEWDATA'); // create a 'NEWDATA' sheet to store imported data
    // loop through csv data array and insert (append) as rows into 'NEWDATA' sheet
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      dataSheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    // rename the report.csv file so it is not processed on next scheduled run
    file.setName("reconciliation data imported on "+(new Date().toString())+".csv");
    SpreadsheetApp.flush();
    CopyWeeklyReconciliationReport();
  }
};

function CopyWeeklyReconciliationReport() {
  var ReportDestFolder = DriveApp.getFolderById('11jwq_eL5pOtspMN9WDUmxJJ3uUuKUex1'); 
  var ReportTemplate = DriveApp.getFileById('1_EIJDgkDtxuEqfQFqUMQpF3oMRf8ng4INwBmJAvdXJA');
  ReportTemplate.makeCopy("NatGeo Weekly Reconciliation Report - "+(new Date().toString()), ReportDestFolder);
};


// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to COMMA.
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];

    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
};