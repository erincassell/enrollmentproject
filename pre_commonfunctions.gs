function getDataValues(name) {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var data = sa.getSheetByName(name).getDataRange().getValues();
  return data;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Regenerate Students')
    .addItem('Get New Students', 'putRegistrations')
    .addToUi();
}

function checkDefined(x) {
  if(typeof(x) == 'undefined') {
    return("");
  } else {
    return x;
  }
}

function refreshSheets() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [];
  sheets.push(sa.getSheetByName("Form Data"));
  
  for(var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var formula = sheet.getRange(1, 1).getFormula();
    sheet.clear();
    sheet.getRange(1, 1).setFormula(formula);
  }
}