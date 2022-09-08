

function getSheetsFromFolder() {
  var date = "09-08"
  // gets the folder from google drive
  const fldr = DriveApp.getFolderById("160gR_xnzNnhH4iUSMuH6pHsDXMnQ_kOI");
  Logger.log("got folder");
  // gets the xlsx files in the specific folder
  const xlsx_files = fldr.getFiles();
  Logger.log("got xlsx files");
  // keeps track of what file we are on
  var count = 0;
  // loop throough every xlsx file in folder
  while (xlsx_files.hasNext()) {
    count = count + 1;
    Logger.log("NEW FILE #" + count);
    // sleep
    Utilities.sleep(1000);
    // find the next xlsx file 
    var xlsx = xlsx_files.next();
    // get the xlsx file and id
    var id = xlsx.getId();
    // copies data in xlsx file to copy to new google sheet file 
    var blob = xlsx.getBlob();
    // sleep
    Utilities.sleep(1000);
    var name = date + "_" + "qcall_" + count;
    // creates new google sheet file to convert xlsx to google sheets
    var newFile = {
      title: name,
      parents: [{ id: "160gR_xnzNnhH4iUSMuH6pHsDXMnQ_kOI" }],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    Logger.log("new SS file created");

    // copies the xlsx content to the new file 
    Drive.Files.insert(newFile, blob);
    // deletes the xlsx file 
    Drive.Files.remove(id);
    Logger.log("old files deleted")
    Logger.log("done")
    // sleep 
    Utilities.sleep(1000);
  }
  combineToSingleSS(date);
}


function combineToSingleSS(date) {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("160gR_xnzNnhH4iUSMuH6pHsDXMnQ_kOI");
  Logger.log("got folder");
  // gets the converted google sheet files in the specific folder
  const files = fldr.getFiles();
  Logger.log("got files");
  Logger.log(files);
  // creates new spreadsheet for data in individual xlsx sheets to be uploaded to
  var combinedSS = SpreadsheetApp.create("Q CALLS " + date);
  // gets new spreadsheet id 
  var id = combinedSS.getId();
  Logger.log(combinedSS.getName());
  // opens new spreadsheet so it can be edited
  var newSS = SpreadsheetApp.openById(id);
  // keeps track of what file we are on
  var count = 0;
  // loops through GS files in folder to add to new spreadsheet
  while (files.hasNext()) {
    count = count + 1;
    // sleep
    Utilities.sleep(1000);
    // gets next file
    var file = files.next();
    Logger.log(file);
    // gets next file name 
    var file_name = file.getName();
    Logger.log(file_name)
    // opens source spreadsheet
    var sh = SpreadsheetApp.openByUrl(file.getUrl());
    // gets sheet in that spreadsheets
    var target_sh = sh.getSheets();
    var source = sortArray(target_sh)
    // loop to copy data from 
    for (var s = 0 in source) {
      // sleep
      Utilities.sleep(1000);
      // gets individual sheet in raw data spreadsheet by index in array
      var sheetR = sh.getSheetByName(source[s]);
      Logger.log(sheetR.getName());
      // copies the sheet from the source spreadsheet to the combined spreadsheeyt
      sheetR.copyTo(newSS).setName(sheetR.getName() + count);
    }
  }
  addSheet(id)
}

function addSheet(id) {
  // opens new spreadsheet so it can be edited
  var newSS = SpreadsheetApp.openById(id);
  // adds sheet to insert data query
  var sheet = newSS.insertSheet("DATA");
  Logger.log("added DATA sheet")
  // gets the first cell 
  var cell = sheet.getRange('A1');
  // add formula to first box 
  cell.setFormula("=QUERY({'Users (2)'!A1:I;'Users (5)'!A2:I;'Users (4)'!A2:I;'Users (3)'!A2:I;'Users (1)'!A2:I;Users!A2:I},'select* where Col1 is not null')")
  Logger.log("queried")
  addQCalls(id)
}

function sortArray(allsheets) {
  // creates empty array to store sheet names
  var sheetNameArray = [];
  // loop through sheets and add each name to array
  for (var i = 0; i < allsheets.length; i++) {
    sheetNameArray.push(allsheets[i].getName());
  }
  // sort the array 
  sheetNameArray.sort(function (a, b) {
    return a.localeCompare(b);
  });
  Logger.log(sheetNameArray);
  // return sorted array
  return sheetNameArray;
}






function addQCalls(id) {
  id = '132lrtFU6vdCw_2dopv13Amv7aO7_omSqntQyFYZhueM';
  // calls the function to get day of week from the date 
  var day = 'WEDNESDAY';
  Logger.log(day.slice(0,2));
  var ss = SpreadsheetApp.openById("1b-mq3_5_Od5a56ivR1KqoHZvT24qGe2gBVA2k-sIB2w");
  var hiddenSheet = ss.getSheetByName("444" + day.slice(0, 3));
  Logger.log(hiddenSheet.getName());

  var ssRD = SpreadsheetApp.openById(id);
  // gets individual sheet with data
  var dataSheet = ssRD.getSheetByName("DATA");
  Logger.log("got DATA");
  // gets range of data to transfer
  var last_col = dataSheet.getLastColumn();
  var last_row = dataSheet.getLastRow();
  var range_input = dataSheet.getRange(1, 1, last_row, last_col);
  Logger.log(last_col);
  Logger.log(last_row);
  // sleep
  Utilities.sleep(1000);
  // copies the values in the the raw data spreadsheet in the range 
  var values = range_input.getValues();
  // sleep
  Utilities.sleep(1000);
  //puts data in sales report
  hiddenSheet.getRange(1, 1, last_row, last_col).setValues(values);


}









