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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Milage Project Tools')
    .addItem('#1 Copy Visible Rows', 'copyVisibleRows')
    .addItem('#2 Reformat Multi-Line Rows', 'reformatMultiLineRowsToNewSheet')
    .addItem('#3 Highlight Matching Rows', 'highlightMatchingRows')
    .addItem('#4 Delete Matching Rows', 'deleteMatchingRows')
    .addToUi();
}