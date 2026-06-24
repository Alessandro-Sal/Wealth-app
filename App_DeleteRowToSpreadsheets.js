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
  
  if (!creditSheet) throw new Error("Foglio Active_Credits mancante.");

  const cData = creditSheet.getDataRange().getValues();
  let rowsToDelete = [];
  let expYear = new Date().getFullYear(); // Valore di default

  // 1. Raccoglie le righe da eliminare e cerca l'anno originale
  if (linkedTxId) {
    for (let i = cData.length - 1; i >= 1; i--) {
      if (String(cData[i][6]).trim() === String(linkedTxId).trim()) {
        rowsToDelete.push(i + 1);
        
        // Estrai l'anno dalla data del credito (colonna B / indice 1)
        let d = new Date(cData[i][1]);
        if (!isNaN(d.getTime())) {
          expYear = d.getFullYear();
        }
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

  // 2. Trova ed elimina la spesa madre nel foglio Expenses in modo dinamico usando l'anno corretto
  if (linkedTxId) {
    const expSheetName = "Expenses Tracker " + expYear;
    const expSheet = ss.getSheetByName(expSheetName);
    
    if (expSheet) {
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
    } else {
      console.log("Foglio Expenses non trovato per l'anno: " + expYear);
    }
  }
  
  return "Success";
}