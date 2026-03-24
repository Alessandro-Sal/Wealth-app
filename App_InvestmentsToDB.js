/**
 * Adds a new investment transaction to the History sheets or Real Assets DB.
 * If the action is a "Withdrawal" (Cash Out), it automatically syncs the entry 
 * to the current year's "Expenses Tracker" as an Investment expense/transfer.
 * Generates a unique Transaction ID to allow synchronized deletion later.
 * * @param {Object} data - The transaction data object.
 * @param {string} data.type - "Crypto", "Stocks", "RealEstate", or "Bond".
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

  // --- NEW: REAL ASSETS & BONDS (Write to DB_RealAssets & DB_Debts) ---
  if (data.type === "RealEstate" || data.type === "Bond") {
    const dbAssets = ss.getSheetByName("DB_RealAssets");
    if (!dbAssets) return "Error: DB_RealAssets sheet missing. Please create it.";
    
    let assetId = (data.type === "RealEstate" ? "RE-" : "BND-") + new Date().getTime().toString().slice(-6);
    let linkedDebtId = "";
    
    // If Real Estate has an associated mortgage, create the Debt first
    if (data.type === "RealEstate" && parseFloat(data.loanAmt) > 0) {
      const dbDebts = ss.getSheetByName("DB_Debts");
      if (dbDebts) {
        linkedDebtId = "DBT-" + new Date().getTime().toString().slice(-6);
        dbDebts.appendRow([
          linkedDebtId, 
          "Mutuo Ipotecario", 
          data.startDate || data.date, 
          parseFloat(data.loanAmt), 
          parseFloat(data.rate), 
          parseFloat(data.years), 
          "", // Monthly payment will be calculated/handled outside
          "Active"
        ]);
      }
    }

    // Prepare row for DB_RealAssets (14 columns total based on our schema)
    let rowData = new Array(14).fill("");
    rowData[0] = assetId;
    rowData[1] = data.type === "RealEstate" ? "Real Estate" : (data.tax === "true" ? "Government Bond" : "Corporate Bond");
    rowData[2] = data.name;
    rowData[3] = data.type === "Bond" ? data.isin : ""; // ISIN
    rowData[4] = data.startDate || data.date; // Purchase Date
    rowData[11] = "Active"; // Status

    if (data.type === "RealEstate") {
      rowData[5] = parseFloat(data.marketVal); // Purchase Price
      rowData[10] = linkedDebtId; // Linked Debt ID
    } else if (data.type === "Bond") {
      let nominal = parseFloat(data.nominal);
      let price = parseFloat(data.price);
      
      rowData[5] = (nominal * price) / 100; // Purchase Price total cost
      rowData[6] = nominal; // Nominal Value
      rowData[7] = parseFloat(data.coupon); // Coupon Rate
      rowData[9] = data.tax === "true" ? "WhiteList" : "Standard"; // Tax Status
      rowData[12] = data.isin; // Ticker Yahoo (ISIN used as fallback ticker)
    }
    
    dbAssets.appendRow(rowData);
    return data.type === "RealEstate" ? "Real Estate & Mortgage Added!" : "Bond Added!";
  }

  
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

  // --- OPTIMIZED BATCH WRITE TO HISTORY ---
  let histRowData = new Array(13).fill("");
  
  histRowData[0] = new Date(data.date);
  histRowData[1] = finalTicker;
  histRowData[2] = data.action;
  histRowData[3] = quantity;

  if (isCrypto) {
      histRowData[4] = totalCostEur;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId; // M = 13
  } else {
      let assetClass = (data.action === "Withdrawal") ? "x" : ((data.type === "Stocks") ? "Stock" : "ETF");
      histRowData[4] = assetClass;
      histRowData[5] = unitPrice;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId; // M = 13
  }

  // Scrive tutta la riga in un colpo solo
  sheet.getRange(newRow, 1, 1, 13).setValues([histRowData]);

  // --- PART 2: EXPENSES TRACKER (Withdrawals Only) ---
  if (data.action === "Withdrawal") {
    const txDate = new Date(data.date);
    const year = txDate.getFullYear();
    const expSheetName = "Expenses Tracker " + year;
    const expSheet = ss.getSheetByName(expSheetName);
    
    if (expSheet && data.bankCol) {
      const startRowExp = 20;
      
      // Find empty row logic
      const colB = expSheet.getRange(startRowExp, 2, Math.max(1, expSheet.getLastRow() - startRowExp + 2), 1).getValues();
      let targetRow = startRowExp;
      for (let i = 0; i < colB.length; i++) {
        if (!colB[i][0]) { targetRow = startRowExp + i; break; }
      }

      const category = isCrypto ? "Crypto" : "Stocks";
      const userNote = (data.note && data.note.trim() !== "") ? data.note : "";

      // Dynamically find "Controllo Automatismi" column
      const headerRowExp = 2;
      const maxCols = expSheet.getLastColumn();
      const headersExp = expSheet.getRange(headerRowExp, 1, 1, maxCols).getValues()[0];
      const syncColIndexExp = headersExp.indexOf("Controllo Automatismi") + 1; 
      
      // OPTIMIZED BATCH WRITE FOR EXPENSES
      // Create an array large enough to hold all data up to the sync column
      const numCols = Math.max(maxCols, syncColIndexExp);
      let expRowData = new Array(numCols).fill(""); 
      
      expRowData[0] = txDate;         // A: Date
      expRowData[1] = 'Investment';   // B: Type
      expRowData[2] = category;       // C: Category
      expRowData[3] = userNote;       // D: Note
      
      const colIndex = parseInt(data.bankCol);
      if (!isNaN(colIndex) && colIndex > 0) {
        expRowData[colIndex - 1] = Math.abs(totalCostEur); // -1 because Array is 0-indexed
      }
      
      if (syncColIndexExp > 0) {
        expRowData[syncColIndexExp - 1] = transactionId;
      }

      // Write everything in a single API call
      expSheet.getRange(targetRow, 1, 1, numCols).setValues([expRowData]);
    }
  }
  return "Investment Saved (" + data.type + ")";
}