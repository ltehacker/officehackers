// Script 1: Highlight rows by directly setting background color based on user input
function highlightMatchingRows() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
  
    // Prompt for the column to search in
    var colResponse = ui.prompt(
      'Column to Search',
      'Enter the letter of the column to search in (e.g., C):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (colResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Script cancelled.');
      return;
    }
  
    var colLetter = colResponse.getResponseText().trim().toUpperCase();
    if (!colLetter.match(/^[A-Z]+$/)) {
      ui.alert('Invalid input. Please enter a valid column letter (e.g., C).');
      return;
    }
  
    var colIndex = columnLetterToIndex(colLetter);
    var numColumns = sheet.getDataRange().getNumColumns();
    if (colIndex < 0 || colIndex >= numColumns) {
      ui.alert('Column letter is out of range for this sheet.');
      return;
    }
  
    // Prompt for the string to search for
    var searchResponse = ui.prompt(
      'Search String',
      'Enter the string to search for (case-insensitive):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (searchResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Script cancelled.');
      return;
    }
  
    var searchString = searchResponse.getResponseText().trim().toLowerCase();
    if (searchString === '') {
      ui.alert('Please enter a non-empty search string.');
      return;
    }
  
    // Get the data
    var range = sheet.getDataRange();
    var values = range.getValues();
    var rowsToHighlight = [];
  
    // Find rows where the specified column matches the search string
    for (var i = 1; i < values.length; i++) { // Start from row 2 (skip header)
      var cellValue = values[i][colIndex];
      if (typeof cellValue === 'string' && cellValue.toLowerCase().includes(searchString)) {
        rowsToHighlight.push(i + 1); // Row numbers are 1-based in Sheets
      }
    }
  
    // Highlight matching rows by setting the background color
    if (rowsToHighlight.length > 0) {
      for (var j = 0; j < rowsToHighlight.length; j++) {
        var rowRange = sheet.getRange(rowsToHighlight[j], 1, 1, numColumns);
        rowRange.setBackground('#FF9999'); // Light red background
      }
      ui.alert(rowsToHighlight.length + ' row(s) highlighted with a light red background.');
    } else {
      ui.alert('No rows found matching the search string.');
    }
  }
  
  // Script 2: Delete rows based on user input (unchanged)
  function deleteMatchingRows() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
  
    // Prompt for the column to search in
    var colResponse = ui.prompt(
      'Column to Search',
      'Enter the letter of the column to search in (e.g., C):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (colResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Script cancelled.');
      return;
    }
  
    var colLetter = colResponse.getResponseText().trim().toUpperCase();
    if (!colLetter.match(/^[A-Z]+$/)) {
      ui.alert('Invalid input. Please enter a valid column letter (e.g., C).');
      return;
    }
  
    var colIndex = columnLetterToIndex(colLetter);
    var numColumns = sheet.getDataRange().getNumColumns();
    if (colIndex < 0 || colIndex >= numColumns) {
      ui.alert('Column letter is out of range for this sheet.');
      return;
    }
  
    // Prompt for the string to search for
    var searchResponse = ui.prompt(
      'Search String',
      'Enter the string to search for (case-insensitive):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (searchResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Script cancelled.');
      return;
    }
  
    var searchString = searchResponse.getResponseText().trim().toLowerCase();
    if (searchString === '') {
      ui.alert('Please enter a non-empty search string.');
      return;
    }
  
    // Get the data
    var range = sheet.getDataRange();
    var values = range.getValues();
    var rowsToDelete = [];
  
    // Find rows where the specified column matches the search string
    for (var i = values.length - 1; i >= 1; i--) { // Start from the bottom to avoid index issues when deleting
      var cellValue = values[i][colIndex];
      if (typeof cellValue === 'string' && cellValue.toLowerCase().includes(searchString)) {
        rowsToDelete.push(i + 1); // Row numbers are 1-based in Sheets
      }
    }
  
    // Delete the matching rows
    if (rowsToDelete.length > 0) {
      var confirmation = ui.alert(
        'Confirm Deletion',
        'Found ' + rowsToDelete.length + ' row(s) to delete. Proceed?',
        ui.ButtonSet.YES_NO
      );
  
      if (confirmation === ui.Button.YES) {
        // Sort rows in descending order to delete from bottom up (avoids shifting issues)
        rowsToDelete.sort((a, b) => b - a);
        for (var j = 0; j < rowsToDelete.length; j++) {
          sheet.deleteRow(rowsToDelete[j]);
        }
        ui.alert(rowsToDelete.length + ' row(s) deleted.');
      } else {
        ui.alert('Deletion cancelled.');
      }
    } else {
      ui.alert('No rows found matching the search string.');
    }
  }
  
  // Helper function to convert column letter to 0-based index
  function columnLetterToIndex(letter) {
    var column = 0;
    for (var i = 0; i < letter.length; i++) {
      column *= 26;
      column += letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return column - 1; // Convert to 0-based index
  }
  
  // Helper function to convert 0-based index to column letter
  function columnIndexToLetter(index) {
    var letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }
  
  // Add a custom menu to run the scripts from the spreadsheet
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Tools')
      .addItem('Highlight Matching Rows', 'highlightMatchingRows')
      .addItem('Delete Matching Rows', 'deleteMatchingRows')
      .addToUi();
  }