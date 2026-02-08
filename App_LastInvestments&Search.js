/**
 * Retrieves the 5 most recent transactions for Stocks and Crypto.
 * Used to populate the "Recent Activity" widget on the dashboard.
 * * @return {Object} An object containing two arrays: 'stocks' and 'crypto'.
 */
function getLastInvestments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Helper to fetch and map data from a specific sheet
  const fetchLast5 = (sheetName, type) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return []; // No data available
    
    // Read from Row 3, first 8 columns
    const dataRange = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
    let transactions = [];

    for (let i = 0; i < dataRange.length; i++) {
      // Check if Date (Col 0) and Ticker (Col 1) exist
      if (dataRange[i][0] && dataRange[i][1]) {
        let dateObj = new Date(dataRange[i][0]);
        let valEur = 0; 
        let valFx = 0;
        
        // Handle Column differences between Stocks and Crypto sheets
        if (type === 'Stocks') {
           // Stocks: EUR in Col 5 (Index), FX in Col 6
           valEur = parseFloat(dataRange[i][5]) || 0; 
           valFx = parseFloat(dataRange[i][6]) || 0;
        } else {
           // Crypto: EUR in Col 4 (Index), FX in Col 5
           valEur = parseFloat(dataRange[i][4]) || 0; 
           valFx = parseFloat(dataRange[i][5]) || 0;
        }

        transactions.push({
          row: 3 + i, 
          sheet: sheetName, 
          date: Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "dd/MM"),
          ticker: dataRange[i][1], 
          action: dataRange[i][2], 
          qty: parseFloat(dataRange[i][3]), 
          eur: valEur, 
          fx: valFx
        });
      }
    }
    // Return last 5, reversed to show newest first
    return transactions.slice(-5).reverse();
  };

  return { 
    stocks: fetchLast5("History B/S Stocks", "Stocks"), 
    crypto: fetchLast5("History B/S Crypto", "Crypto") 
  };
}

/**
 * Search engine for the Investment History.
 * Filters transactions by Ticker (partial match), Asset Type, and Year.
 * * @param {string} query - The ticker symbol to search for (e.g., "AAPL").
 * @param {string} assetType - Filter by "Stocks", "Crypto", or "" (All).
 * @param {string|number} yearFilter - Filter by specific year.
 * @return {Array<Object>} List of matching transactions, sorted by newest first.
 */
function searchInvestHistory(query, assetType, yearFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let results = [];
  const lowerQuery = query ? query.toLowerCase() : "";

  // Helper to scan a sheet with specific column mapping
  const scanSheet = (sheetName, typeLabel, colEurIdx, colForexIdx) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return; 
    
    const data = sheet.getRange(3, 1, lastRow - 2, 8).getValues();

    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue; 
      
      let rowDate = new Date(data[i][0]);
      let ticker = String(data[i][1]); 
      let action = String(data[i][2]); 
      let qty = data[i][3];            
      
      // Apply Year Filter
      if (yearFilter && yearFilter !== "" && rowDate.getFullYear() != yearFilter) continue;
      
      // Apply Query Filter (Ticker search)
      if (lowerQuery !== "") { 
        if (!ticker.toLowerCase().includes(lowerQuery)) continue; 
      }
      
      let valEur = parseFloat(data[i][colEurIdx]) || 0;
      let valForex = parseFloat(data[i][colForexIdx]) || 0;
      
      results.push({
        date: Utilities.formatDate(rowDate, ss.getSpreadsheetTimeZone(), "dd/MM/yy"),
        ticker: ticker, 
        action: action, 
        qty: qty, 
        eur: valEur, 
        forex: valForex, 
        type: typeLabel
      });
    }
  };

  // Execution based on Asset Type selection
  if (assetType === "" || assetType === "Stocks") {
    scanSheet("History B/S Stocks", "Stocks", 5, 6); // Cols F & G
  }
  if (assetType === "" || assetType === "Crypto") {
    scanSheet("History B/S Crypto", "Crypto", 4, 5); // Cols E & F
  }

  return results.reverse(); // Newest first
}