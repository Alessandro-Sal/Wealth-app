/**
 * "Fill Down" Function: Fills missing dates by copying them from the cell above.
 * Iterates row by row to plug gaps in the date column.
 * 
 */
function fillDatesCascade() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = "Expenses Tracker 2021"; 
  const SHEET = SS.getSheetByName(SHEET_NAME);
  
  if (!SHEET) {
    SpreadsheetApp.getUi().alert("Sheet '" + SHEET_NAME + "' not found!");
    return;
  }

  // --- SETTINGS ---
  // If your archive starts at row 20 (as per previous logic), keep 20.
  // If it starts at the top, change to 2.
  const START_ROW = 20; 
  
  // The column index for dates. Column A = 1.
  const DATE_COLUMN = 1; 
  
  const lastRow = SHEET.getLastRow();

  if (lastRow < START_ROW) {
    SpreadsheetApp.getUi().alert("Not enough data to run the script.");
    return;
  }

  // Read all column data in one batch (Performance Optimization)
  const numRows = lastRow - START_ROW + 1;
  const range = SHEET.getRange(START_ROW, DATE_COLUMN, numRows, 1);
  let values = range.getValues(); 

  let lastValidDate = null;
  let changesMade = false;

  // --- ROW BY ROW LOOP ---
  for (let i = 0; i < values.length; i++) {
    let cellValue = values[i][0];

    // CASE 1: Cell has a date
    if (cellValue !== "" && cellValue !== null) {
      lastValidDate = cellValue; // Update the current "active" date
    } 
    // CASE 2: Cell is empty, but we have a date in memory
    else if (lastValidDate !== null) {
      values[i][0] = lastValidDate; // Paste the stored date
      changesMade = true;
    }
  }

  // Write everything back to the sheet only at the end (Batch Write)
  if (changesMade) {
    range.setValues(values);
    SpreadsheetApp.getUi().alert("Dates filled successfully!");
  } else {
    SpreadsheetApp.getUi().alert("All rows already had a date. No changes made.");
  }
}