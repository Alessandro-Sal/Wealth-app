/**
 * Deletes a specific row from a sheet given its index.
 * Implements bidirectional synchronization:
 * 1. If deleting from "Expenses Tracker", it also removes the corresponding entry in "History" sheets.
 * 2. If deleting from "History", it removes the entry from the current year's "Expenses Tracker".
 * * @param {string} sheetName - The name of the sheet to delete from.
 * @param {number} rowIndex - The 1-based index of the row to delete.
 * @return {string} Status message ("Deleted" or Error).
 */
function deleteRow(sheetName, rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return "Error: Sheet not found";

  // --- 1. RETRIEVE ID BEFORE DELETION ---
  let idToDelete = "";
  let isFromExpenses = sheetName.startsWith("Expenses Tracker");
  let isFromHistory = sheetName.includes("History B/S");

  try {
    // If in Expenses, ID is in Column 35 (AI)
    if (isFromExpenses) {
      idToDelete = sheet.getRange(rowIndex, 35).getValue();
    } 
    // If in History, ID is in Column 13 (M)
    else if (isFromHistory) {
      idToDelete = sheet.getRange(rowIndex, 13).getValue();
    }
  } catch (e) {
    console.log("No ID found or read error: " + e);
  }

  // --- 2. SYNCHRONIZED DELETION ---
  if (idToDelete && String(idToDelete).trim().startsWith("ID_")) {
    const id = String(idToDelete).trim();

    // CASE A: Deleting from EXPENSES -> Sync delete in History sheets
    if (isFromExpenses) {
      findAndDeleteById(ss, "History B/S Stocks", id, 13); // 13 = Col M
      findAndDeleteById(ss, "History B/S Crypto", id, 13);
    }
    
    // CASE B: Deleting from HISTORY -> Sync delete in Expenses
    else if (isFromHistory) {
      // Determine which year to search. 
      // Defaults to checking the current year's sheet.
      const currentYear = new Date().getFullYear();
      const expSheetName = "Expenses Tracker " + currentYear;
      
      // Note: To support past years, logic should be extended here.
      findAndDeleteById(ss, expSheetName, id, 35); // 35 = Col AI
    }
  }

  // --- 3. DELETE ORIGINAL ROW ---
  sheet.deleteRow(rowIndex);
  return "Deleted";
}

/**
 * Helper function to locate and delete a row by its unique ID in a target sheet.
 * * @param {Spreadsheet} ss - The active spreadsheet object.
 * @param {string} targetSheetName - The name of the sheet to search in.
 * @param {string} id - The unique ID string to match.
 * @param {number} colIndex - The column index where the ID is stored.
 */
function findAndDeleteById(ss, targetSheetName, id, colIndex) {
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) return;

  const lastRow = targetSheet.getLastRow();
  
  // Expenses sheet data starts at row 20; History usually starts at row 1 or 3.
  const startRow = targetSheetName.startsWith("Expenses") ? 20 : 1;
  
  if (lastRow < startRow) return;

  // Read only the ID column to optimize performance
  const range = targetSheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1);
  const values = range.getValues();

  // Reverse loop to safely delete rows without messing up indices
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === id) {
      // Found match! Delete the actual row
      // Index 'i' is relative to the range (0-based). Actual Row = startRow + i
      targetSheet.deleteRow(startRow + i);
      console.log("Sync Delete: Removed linked row in " + targetSheetName);
      
      // Break after first match (IDs are assumed unique)
      break; 
    }
  }
}