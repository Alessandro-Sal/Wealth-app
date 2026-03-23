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
  let expSyncCol = 35; // Fallback

  try {
    // Seleziona la colonna dinamicamente se siamo in Expenses (intestazione alla riga 2)
    if (isFromExpenses) {
      const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      expSyncCol = headers.indexOf("Controllo Automatismi") + 1;
      
      if (expSyncCol > 0) {
        idToDelete = sheet.getRange(rowIndex, expSyncCol).getValue();
      }
    } 
    // If in History, ID is always in Column 13 (M)
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
      const currentYear = new Date().getFullYear();
      const expSheetName = "Expenses Tracker " + currentYear;
      const targetExpSheet = ss.getSheetByName(expSheetName);
      
      if (targetExpSheet) {
        // Find dynamic column for the target expenses sheet (intestazione alla riga 2)
        const headers = targetExpSheet.getRange(2, 1, 1, targetExpSheet.getLastColumn()).getValues()[0];
        const targetExpCol = headers.indexOf("Controllo Automatismi") + 1;
        
        if (targetExpCol > 0) {
          findAndDeleteById(ss, expSheetName, id, targetExpCol);
        }
      }
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
  const startRow = targetSheetName.startsWith("Expenses") ? 20 : 1;
  
  if (lastRow < startRow) return;

  // Read only the ID column to optimize performance
  const range = targetSheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1);
  const values = range.getValues();

  // Reverse loop to safely delete rows without messing up indices
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === id) {
      targetSheet.deleteRow(startRow + i);
      console.log("Sync Delete: Removed linked row in " + targetSheetName);
      break; 
    }
  }
}

/**
 * Rimuove TUTTI i crediti di uno split e la transazione originale associata.
 */
function removeCreditAndOriginalTx(creditId, linkedTxId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creditSheet = ss.getSheetByName("Active_Credits");
  const currentYear = new Date().getFullYear();
  const expSheetName = "Expenses Tracker " + currentYear;
  const expSheet = ss.getSheetByName(expSheetName);

  if (!creditSheet) throw new Error("Foglio Active_Credits mancante.");

  const cData = creditSheet.getDataRange().getValues();
  let rowsToDelete = [];

  // 1. Raccoglie le righe da eliminare (ottimizzato per bulk virtuale)
  if (linkedTxId) {
    for (let i = cData.length - 1; i >= 1; i--) {
      if (String(cData[i][6]).trim() === String(linkedTxId).trim()) {
        rowsToDelete.push(i + 1);
      }
    }
  } else {
    for (let i = cData.length - 1; i >= 1; i--) {
      if (String(cData[i][0]) === String(creditId)) {
        rowsToDelete.push(i + 1);
        break;
      }
    }
  }

  // Elimina le righe dei crediti (dal basso verso l'alto per sicurezza)
  rowsToDelete.forEach(r => creditSheet.deleteRow(r));

  // 2. Trova ed elimina la spesa madre nel foglio Expenses in modo dinamico
  if (linkedTxId && expSheet) {
    // Intestazione alla riga 2
    const headers = expSheet.getRange(2, 1, 1, expSheet.getLastColumn()).getValues()[0];
    const syncColIndex = headers.indexOf("Controllo Automatismi") + 1;

    if (syncColIndex > 0) {
      const eData = expSheet.getRange(20, syncColIndex, expSheet.getLastRow() - 19, 1).getValues(); 
      for (let i = eData.length - 1; i >= 0; i--) {
        if (String(eData[i][0]).trim() === String(linkedTxId).trim()) {
          deleteRow(expSheetName, i + 20); 
          break; 
        }
      }
    }
  }
  
  return "Success";
}