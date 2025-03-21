function reformatMultiLineRowsToNewSheet() {
  // Get the active spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Prompt the user for the first column with multi-line data
  var response1 = ui.prompt(
    'First Multi-Line Column',
    'Enter the letter of the first column with multi-line data (e.g., C for Title):',
    ui.ButtonSet.OK_CANCEL
  );

  // Check if the user clicked "Cancel" or closed the prompt
  if (response1.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Script cancelled.');
    return;
  }

  var colLetter1 = response1.getResponseText().trim().toUpperCase();
  if (!colLetter1.match(/^[A-Z]+$/)) {
    ui.alert('Invalid input. Please enter a valid column letter (e.g., C).');
    return;
  }

  // Prompt the user for the second column with multi-line data
  var response2 = ui.prompt(
    'Second Multi-Line Column',
    'Enter the letter of the second column with multi-line data (e.g., G for Destinations):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response2.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Script cancelled.');
    return;
  }

  var colLetter2 = response2.getResponseText().trim().toUpperCase();
  if (!colLetter2.match(/^[A-Z]+$/)) {
    ui.alert('Invalid input. Please enter a valid column letter (e.g., G).');
    return;
  }

  // Convert column letters to 0-based indices
  var multiLineCol1 = columnLetterToIndex(colLetter1);
  var multiLineCol2 = columnLetterToIndex(colLetter2);

  // Validate the indices
  var numColumns = sourceSheet.getDataRange().getNumColumns();
  if (multiLineCol1 < 0 || multiLineCol1 >= numColumns || multiLineCol2 < 0 || multiLineCol2 >= numColumns) {
    ui.alert('One or both column letters are out of range for this sheet.');
    return;
  }

  // Get the data
  var range = sourceSheet.getDataRange();
  var values = range.getValues();
  var headers = values[0]; // First row is headers
  var data = values.slice(1); // Data starts from second row
  var newData = [headers]; // Start with headers for the new sheet

  // Loop through each row of data
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var lines1 = [];
    var lines2 = [];

    // Split the multi-line columns into arrays of lines
    if (typeof row[multiLineCol1] === 'string' && row[multiLineCol1].includes('\n')) {
      lines1 = row[multiLineCol1].split('\n').filter(line => line.trim() !== '');
    } else {
      lines1 = [row[multiLineCol1]]; // Treat as a single line
    }

    if (typeof row[multiLineCol2] === 'string' && row[multiLineCol2].includes('\n')) {
      lines2 = row[multiLineCol2].split('\n').filter(line => line.trim() !== '');
    } else {
      lines2 = [row[multiLineCol2]]; // Treat as a single line
    }

    // Check for mismatches in line counts (for debugging)
    if (lines1.length !== lines2.length) {
      ui.alert(
        'Mismatch in row ' + (i + 2) + ': ' +
        'Column ' + colLetter1 + ' has ' + lines1.length + ' lines, ' +
        'Column ' + colLetter2 + ' has ' + lines2.length + ' lines. ' +
        'Please check the data.'
      );
      return; // Stop the script if a mismatch is found
    }

    // Create new rows by pairing the lines
    for (var j = 0; j < lines1.length; j++) {
      var newRow = row.slice(); // Copy the original row
      newRow[multiLineCol1] = lines1[j]; // Set the j-th line from the first multi-line column
      newRow[multiLineCol2] = lines2[j]; // Set the j-th line from the second multi-line column
      newData.push(newRow);
    }
  }

  // Create a new sheet for the reformatted data with a unique name
  var baseName = sourceSheet.getName() + ' SPLIT';
  var newSheetName = baseName;
  var counter = 1;
  while (spreadsheet.getSheetByName(newSheetName)) {
    newSheetName = baseName + ' ' + counter;
    counter++;
  }
  var newSheet = spreadsheet.insertSheet(newSheetName);

  // Write the new data to the new sheet
  newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
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

// Add a custom menu to run the script from the spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
    .addItem('Reformat Multi-Line Rows', 'reformatMultiLineRowsToNewSheet')
    .addToUi();
}