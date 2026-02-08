/**
 * Retrieves the 5 most recent transactions from the current expense tracker.
 * Used for the "Recent Transactions" widget on the dashboard.
 * * @return {Array<Object>} List of the last 5 transactions with date, type, category, and amount.
 */
function getLastTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Target specific year sheet (Hardcoded to 2026, could be dynamic)
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  
  if (!sheet) return [];
  
  const startRow = 20; 
  const lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) return [];
  
  // Read Data (Cols A to J)
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues();
  let transactions = [];
  
  for (let i = 0; i < dataRange.length; i++) {
    // Check if Type (Col B) exists
    if (dataRange[i][1]) { 
      let totalAmount = 0; 
      // Sum Bank Columns (Indices 4-9 / Cols E-J)
      for (let j = 4; j <= 9; j++) { 
        let val = parseFloat(dataRange[i][j]); 
        if (!isNaN(val)) totalAmount += val; 
      }
      
      transactions.push({ 
        row: startRow + i, 
        date: Utilities.formatDate(new Date(dataRange[i][0]), ss.getSpreadsheetTimeZone(), "dd/MM"), 
        type: dataRange[i][1], 
        category: dataRange[i][2], 
        details: dataRange[i][3], 
        amount: totalAmount 
      });
    }
  }
  
  // Return last 5 items, reversed (Newest first)
  return transactions.slice(-5).reverse();
}

/**
 * Search engine for Expenses. 
 * Supports filtering by text query, type, category, month, and year.
 * Dynamically selects the correct sheet based on the requested year.
 * * @param {string} query - Text search for notes or categories.
 * @param {string} typeFilter - Filter by "Expense", "Income", etc.
 * @param {string} categoryFilter - Filter by specific category.
 * @param {number} monthFilter - Filter by month index (0-11).
 * @param {number} yearFilter - Filter by year (YYYY).
 * @return {Array<Object>} List of matching transactions (max 100).
 */
function searchTransactions(query, typeFilter, categoryFilter, monthFilter, yearFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- YEAR SELECTION LOGIC ---
  // If a year filter is provided, use it to construct the sheet name.
  // Otherwise, default to the current year (2026).
  let targetYear = yearFilter ? String(yearFilter) : "2026";
  let sheetName = "Expenses Tracker " + targetYear; 
  // ----------------------------

  const sheet = ss.getSheetByName(sheetName);
  
  // If the sheet for that year doesn't exist (e.g., "Expenses Tracker 2020"), return empty.
  if (!sheet) return [];

  const startRow = 20; 
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues();
  let results = [];
  const lowerQuery = query ? query.toLowerCase() : "";

  // Loop backwards to show newest results first
  for (let i = dataRange.length - 1; i >= 0; i--) {
    if (!dataRange[i][0]) continue;
    
    let rowDate = new Date(dataRange[i][0]);
    let rowType = dataRange[i][1]; 
    let rowCat = dataRange[i][2] || ""; 
    let rowNote = dataRange[i][3] || "";

    // --- APPLY FILTERS ---
    
    // Month Filter
    if (monthFilter !== "" && rowDate.getMonth() != monthFilter) continue;
    
    // Year Filter: Technically redundant if we selected the right sheet, 
    // but kept as a safety check for bad dates within the sheet.
    if (yearFilter !== "" && rowDate.getFullYear() != yearFilter) continue;
    
    if (typeFilter && typeFilter !== "" && rowType !== typeFilter) continue;
    if (categoryFilter && categoryFilter !== "" && rowCat !== categoryFilter) continue;
    
    // Text Query Filter (Checks both Note and Category)
    if (lowerQuery !== "") { 
      if (!rowNote.toLowerCase().includes(lowerQuery) && !rowCat.toLowerCase().includes(lowerQuery)) continue; 
    }

    // Calculate Amount (Sum of bank columns)
    let totalAmount = 0; 
    for (let j = 4; j <= 9; j++) { 
      let val = parseFloat(dataRange[i][j]); 
      if (!isNaN(val)) totalAmount += val; 
    }

    results.push({ 
      date: Utilities.formatDate(rowDate, ss.getSpreadsheetTimeZone(), "dd/MM/yy"), 
      category: rowCat, 
      details: rowNote, 
      amount: totalAmount, 
      type: rowType 
    });

    // Limit results to 100 for performance
    if (results.length >= 100) break;
  }
  return results;
}