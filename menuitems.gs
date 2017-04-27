function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Regenerate Students')
    .addItem('Get New Students', 'compareandadd')
    .addToUi();
}