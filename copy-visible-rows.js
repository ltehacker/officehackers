function copyVisibleRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Visible Rows');
    
    for (var i = 0; i < values.length; i++) {
      if (!sheet.isRowHiddenByUser(i + 1)) {
        newSheet.appendRow(values[i]);
      }
    }
  }

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Tools')
      .addItem('Copy Visible Rows', 'copyVisibleRows')
      .addToUi();
  }