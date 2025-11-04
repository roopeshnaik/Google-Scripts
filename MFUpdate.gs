function getAMFIData() {  
  var navSheet = SpreadsheetApp.getActive().getSheetByName('NAV').getRange('A1');
  navSheet.clearContent();
  navSheet.setValue('=IMPORTDATA("https://www.amfiindia.com/spages/NAVAll.txt", ";")');
}

function updateMFSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  updateMFSheet("MF_ICICI", "N2:N6");
  //ss.getSheetByName("MF_Zero").activate();
  //SpreadsheetApp.getUi().alert("ICICI MF Sheet Updated");
  showSelfClosingAlert('ICICI MF Sheet Updated.', 1);
  updateMFSheet("MF_Zero", "N2:N7");
  showSelfClosingAlert('Zerodha MF Sheet Updated.', 1);
}

// Get AMFI code and do lookup in NAV sheet
function updateMFSheet(sheetToUpdate, amfiRange) {
  var mfSheet = SpreadsheetApp.getActive().getSheetByName(sheetToUpdate);
  var searchRange = mfSheet.getRange(amfiRange);

  for(var i = 0; i < searchRange.getNumRows(); i++) {
    var list1data = searchRange.getValues()
    vlookupscript(i + 2, parseInt(list1data[i]), sheetToUpdate);
  }
}

// Get data from NAV sheet and update in MF sheets.
function vlookupscript(cellNo, searchValue, sheetToUpdate) {
  var mfSheet = SpreadsheetApp.getActive().getSheetByName(sheetToUpdate);
  var navSheet = SpreadsheetApp.getActive().getSheetByName("NAV").getRange('A:F').getValues();
  var dataList = navSheet.map(x => x[0]);
  var index = dataList.indexOf(searchValue);
  if (index === -1) {
    getAMFIData();
    //mfSheet.getRange('E14').setValue('Error');
  } 
  else {
    var foundValue = navSheet[index][4];
    mfSheet.getRange(cellNo, 5).setValue(foundValue);
    //mfSheet.getRange('E14').setValue('Updated');
  }
}

function showSelfClosingAlert(message, durationSeconds) {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>' +
    '  function closeAlert() {' +
    '    google.script.host.close();' +
    '  }' +
    '  setTimeout(closeAlert, ' + (durationSeconds * 1000) + ');' +
    '</script>' +
    '<p>' + message + '</p>'
  )
  .setWidth(200)
  .setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Status');
}

