/**
 * Adds a new investment transaction to the History sheets.
 * If the action is a "Withdrawal" (Cash Out), it automatically syncs the entry 
 * to the current year's "Expenses Tracker" as an Investment expense/transfer.
 * Generates a unique Transaction ID to allow synchronized deletion later.
 * * * @param {Object} data - The transaction data object.
 * @param {string} data.type - "Crypto" or "Stocks".
 * @param {string} data.action - "Deposit", "Withdrawal", "Buy", "Sell".
 * @param {string} data.ticker - The asset symbol (e.g., "BTC", "AAPL").
 * @param {number} data.qty - Quantity of the asset.
 * @param {number} data.costEur - Total cost in EUR.
 * @param {number} data.costUsd - Total cost in USD.
 * @param {string} [data.bankCol] - The column index of the bank used (required for withdrawals).
 * @param {string} [data.note] - Optional user notes.
 * @return {string} Status message.
 */
function addInvestTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Generate a unique ID to link History and Expenses entries (crucial for synchronized deletion)
  const transactionId = "ID_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);

  // --- PART 1: INVESTMENT HISTORY (History B/S) ---
  const isCrypto = data.type === "Crypto";
  const sheetName = isCrypto ? "History B/S Crypto" : "History B/S Stocks";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet missing (" + sheetName + ")";

  // Find first empty row in History (Starting from Row 3)
  const lastRow = sheet.getLastRow();
  let newRow = lastRow + 1;
  const startRowHist = 3;
  
  if (lastRow >= startRowHist) {
    // Scan Column A to find the first actual empty cell
    const colA = sheet.getRange(startRowHist, 1, lastRow - startRowHist + 2, 1).getValues();
    for (let i = 0; i < colA.length; i++) {
      if (!colA[i][0]) { newRow = startRowHist + i; break; }
    }
  } else {
    newRow = startRowHist;
  }

  const quantity = parseFloat(data.qty) || 0;
  const totalCostEur = parseFloat(data.costEur) || 0;
  const totalCostUsd = parseFloat(data.costUsd) || 0;
  let unitPrice = quantity !== 0 ? totalCostEur / quantity : 0;

  // Ticker Normalization
  let finalTicker = data.ticker.trim();
  if (finalTicker.toLowerCase() === "cash") {
    finalTicker = "Cash"; 
  } else {
    finalTicker = finalTicker.toUpperCase();
  }

  // Write to History Sheet
  sheet.getRange(newRow, 1).setValue(new Date(data.date));
  sheet.getRange(newRow, 2).setValue(finalTicker);
  sheet.getRange(newRow, 3).setValue(data.action);
  sheet.getRange(newRow, 4).setValue(quantity);

  if (isCrypto) {
      sheet.getRange(newRow, 5).setValue(totalCostEur);
      sheet.getRange(newRow, 7).setValue(totalCostUsd);
      // Write ID to Column M (13) for sync
      sheet.getRange(newRow, 13).setValue(transactionId);
  } else {
      let assetClass = "";
      if (data.action === "Withdrawal") { assetClass = "x"; } 
      else { assetClass = (data.type === "Stocks") ? "Stock" : "ETF"; }
      
      sheet.getRange(newRow, 5).setValue(assetClass);
      sheet.getRange(newRow, 6).setValue(unitPrice);
      sheet.getRange(newRow, 7).setValue(totalCostUsd);
      // Write ID to Column M (13) for sync
      sheet.getRange(newRow, 13).setValue(transactionId);
  }

  // --- PART 2: EXPENSES TRACKER (Withdrawals Only) ---
  if (data.action === "Withdrawal") {
    const txDate = new Date(data.date);
    const year = txDate.getFullYear();
    const expSheetName = "Expenses Tracker " + year;
    const expSheet = ss.getSheetByName(expSheetName);
    
    // Ensure sheet exists and a bank column was selected
    if (expSheet && data.bankCol) {
      const startRowExp = 20;
      
      // Find empty row logic for Expenses Sheet
      const colB = expSheet.getRange(startRowExp, 2, Math.max(1, expSheet.getLastRow() - startRowExp + 2), 1).getValues();
      let targetRow = startRowExp;
      for (let i = 0; i < colB.length; i++) {
        if (!colB[i][0]) { 
          targetRow = startRowExp + i;
          break;
        }
      }

      const category = isCrypto ? "Crypto" : "Azioni";
      
      // --- CLEAN NOTES LOGIC ---
      // Only write user input notes. Keep cell empty if no note provided.
      const userNote = (data.note && data.note.trim() !== "") ? data.note : "";

      // Write Expense Row
      expSheet.getRange(targetRow, 1).setValue(txDate);        // A: Date
      expSheet.getRange(targetRow, 2).setValue('Investment');  // B: Type
      expSheet.getRange(targetRow, 3).setValue(category);      // C: Category
      expSheet.getRange(targetRow, 4).setValue(userNote);      // D: Note
      
      // Write Amount to the specific Bank Column
      const colIndex = parseInt(data.bankCol);
      if (!isNaN(colIndex)) {
        expSheet.getRange(targetRow, colIndex).setValue(Math.abs(totalCostEur));
      }

      // --- CRITICAL: SYNC ID WRITING ---
      // Column AI (35) is used for hidden IDs in Expenses Tracker
      expSheet.getRange(targetRow, 35).setValue(transactionId);
    }
  }
  return "Investment Saved (" + data.type + ")";
}