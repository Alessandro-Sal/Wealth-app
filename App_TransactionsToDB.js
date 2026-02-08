/**
 * Main function to record standard transactions (Expenses/Income).
 * Features an auto-sync mechanism that replicates 'Investment' entries into the 
 * respective History sheets (Stocks/Crypto) as 'Cash Deposits', linking them via a unique ID.
 * * @param {Object} data - Transaction data object.
 * @param {string} data.type - Transaction type (Expense, Income, Investment).
 * @param {string} data.category - Category (e.g., "Alimentazione", "Azioni").
 * @param {string} data.details - Description or notes.
 * @param {Object} data.amounts - Key-value pair of { columnIndex: amount }.
 * @return {string} Success message.
 */
function addTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Target Sheet: Expenses Tracker 2026 (Hardcoded year for now)
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  if (!sheet) return "Error: Sheet not found";

  const startRow = 20;
  
  // Logic to find the first truly empty row in the tracker
  const colB = sheet.getRange(startRow, 2, Math.max(1, sheet.getLastRow() - startRow + 1), 1).getValues();
  let newRow = startRow;
  for (let i = 0; i < colB.length; i++) {
    if (colB[i][0] === "" || colB[i][0] === null) {
      newRow = startRow + i;
      break;
    }
    if (i === colB.length - 1) newRow = startRow + i + 1;
  }

  // 1. WRITE TO EXPENSES TRACKER (Standard Operation)
  const dateVal = new Date();
  sheet.getRange(newRow, 1).setValue(dateVal);
  sheet.getRange(newRow, 2).setValue(data.type);
  sheet.getRange(newRow, 3).setValue(data.category);
  sheet.getRange(newRow, 4).setValue(data.details);

  // Write amounts and calculate total for potential investment sync
  let totalInvestAmount = 0;
  for (let col in data.amounts) {
    let val = parseFloat(data.amounts[col]);
    sheet.getRange(newRow, parseInt(col)).setValue(val);
    
    // Sum absolute values if columns are between 5 (E) and 9 (I) - Typical Bank Columns
    if (parseInt(col) >= 5 && parseInt(col) <= 9) {
      totalInvestAmount += Math.abs(val || 0);
    }
  }

  // --- 2. IMMEDIATE SYNCHRONIZATION (AUTO-LINK) ---
  // If Type is "Investment", immediately copy to History as a "Deposit"
  if (data.type === "Investment") {
    
    let destSheetName = null;
    let isCrypto = false;

    // Check Category to determine destination (Strings must match Dropdown values)
    if (data.category === "Azioni") destSheetName = "History B/S Stocks";
    if (data.category === "Crypto") { destSheetName = "History B/S Crypto"; isCrypto = true; }

    if (destSheetName) {
      const destSheet = ss.getSheetByName(destSheetName);
      if (destSheet) {
        // Generate Unique ID based on timestamp and row
        const newId = "ID_" + new Date().getTime() + "_" + newRow;

        // Find empty row in History sheet
        const lastHistRow = destSheet.getLastRow();
        let histRow = 1;
        
        // Search for first free row by checking Column A (Date) backwards
        const histDates = destSheet.getRange("A1:A" + (lastHistRow + 1)).getValues();
        for (let j = histDates.length - 1; j >= 0; j--) {
          if (histDates[j][0] !== "" && histDates[j][0] != null) {
            histRow = j + 2;
            break;
          }
        }
        if (lastHistRow === 0) histRow = 1;

        // Write to History: A=Date, B=Ticker(Cash), C=Action(Deposit), D=Qty(1), E=Class(x), F=Amount
        destSheet.getRange(histRow, 1, 1, 6).setValues([[dateVal, "Cash", "Deposit", 1, "x", totalInvestAmount]]);
        
        // If Crypto, write amount to Col H (8) as well (specific formatting for Crypto sheet)
        if (isCrypto) {
          destSheet.getRange(histRow, 8).setValue(totalInvestAmount);
        }

        // Write ID to History (Col M = 13)
        destSheet.getRange(histRow, 13).setValue(newId);

        // WRITE ID TO EXPENSES (Col AI = 35) - Links the two rows
        sheet.getRange(newRow, 35).setValue(newId);
        
        // Force save to ensure data integrity across sheets
        SpreadsheetApp.flush();
      }
    }
  }
  // ----------------------------------------------------

  return "Saved Successfully";
}