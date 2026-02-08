/**
 * Retrieves the current watchlist data.
 * Reads extended metrics (P/E, Beta, 52w Range) and the live EUR/USD exchange rate from cell Z1.
 * Ensures strict reading of 15 columns to prevent data misalignment.
 * * @return {Array<Object>} List of watchlist items with financial metrics.
 */
function getWatchlistData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Watchlist");
    
    // Return empty if sheet doesn't exist or has no data rows
    if (!sheet || sheet.getLastRow() < 2) return [];

    // 1. Read EUR/USD Exchange Rate (Stored in Cell Z1)
    if (sheet.getRange("Z1").getValue() === "") {
       sheet.getRange("Z1").setFormula('=GOOGLEFINANCE("CURRENCY:EURUSD")');
    }
    const eurUsdRate = sheet.getRange("Z1").getValue() || 1.08; 

    const lastRow = sheet.getLastRow();
    
    // --- CRITICAL FIX: FORCE READING 15 COLUMNS ---
    // Previously used dynamic column counting which caused misalignment if a cell was empty.
    // Now strictly reads columns 1 (A) to 15 (O).
    
    // Safety check: ensure sheet has enough physical columns
    const maxCols = sheet.getMaxColumns();
    const colsToRead = maxCols < 15 ? maxCols : 15;

    const data = sheet.getRange(2, 1, lastRow - 1, colsToRead).getDisplayValues();

    return data.map((row, index) => ({
      row: index + 2,
      ticker: String(row[0] || ""),
      name: String(row[1] || row[0]),
      price: String(row[2] || "0"),
      pct: String(row[3] || "0%"),
      curr: String(row[4] || "EUR"), 
      pe: String(row[5] || "-"),
      eps: String(row[6] || "-"),
      mkt: String(row[7] || "0"),
      vol: String(row[8] || "0"),
      volAvg: String(row[9] || "0"),
      dayH: String(row[10] || "0"),
      dayL: String(row[11] || "0"),
      yearH: String(row[12] || "0"),
      yearL: String(row[13] || "0"),
      beta: String(row[14] || "-"), // If row[14] is undefined, default to "-"
      rate: eurUsdRate
    }));

  } catch (e) {
    console.error("Watchlist Error: " + e.toString());
    return [];
  }
}

/**
 * --- ADD TO WATCHLIST ---
 * Adds a new ticker to the Watchlist sheet and populates it with GOOGLEFINANCE formulas.
 * Checks for duplicates before adding.
 * * @param {string} ticker - The stock ticker symbol (e.g., "AAPL").
 * @return {string} Status message.
 */
function addToWatchlist(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Watchlist");
  
  // Create sheet if missing, with headers
  if (!sheet) {
    sheet = ss.insertSheet("Watchlist");
    sheet.appendRow([
      "Ticker", "Name", "Price", "Change %", "Currency", 
      "P/E", "EPS", "Market Cap", "Volume", "Avg Vol", 
      "Day High", "Day Low", "52w High", "52w Low", "Beta"
    ]);
    // Set Exchange Rate Formula in Z1
    sheet.getRange("Z1").setFormula('=GOOGLEFINANCE("CURRENCY:EURUSD")');
  }

  const cleanTicker = ticker.toUpperCase().trim();
  
  // Check for duplicates
  const data = sheet.getDataRange().getValues();
  for(let i=0; i<data.length; i++) {
    if(String(data[i][0]).toUpperCase() === cleanTicker) return "Ticker already exists!";
  }

  const lastRow = sheet.getLastRow() + 1;
  const c = "A" + lastRow; // Cell reference for the Ticker (e.g., A2)

  // 1. Write Ticker (Plain Text)
  sheet.getRange(lastRow, 1).setValue(cleanTicker);

  // 2. Prepare Formulas for Columns B to O
  // Using IFERROR to handle API failures gracefully
  const formulas = [
    // Array of formula strings
    `=IFERROR(GOOGLEFINANCE(${c}; "name"); "${cleanTicker}")`,      // B: Name
    `=IFERROR(GOOGLEFINANCE(${c}; "price"); 0)`,                    // C: Price
    `=IFERROR(GOOGLEFINANCE(${c}; "changepct")/100; 0)`,            // D: % (Div by 100 for formatting)
    `=IFERROR(GOOGLEFINANCE(${c}; "currency"); "-")`,               // E: Currency
    `=IFERROR(GOOGLEFINANCE(${c}; "pe"); "-")`,                     // F: P/E
    `=IFERROR(GOOGLEFINANCE(${c}; "eps"); "-")`,                    // G: EPS
    `=IFERROR(GOOGLEFINANCE(${c}; "marketcap"); 0)`,                // H: Mkt Cap
    `=IFERROR(GOOGLEFINANCE(${c}; "volume"); 0)`,                   // I: Volume
    `=IFERROR(GOOGLEFINANCE(${c}; "volumeavg"); 0)`,                // J: Avg Vol
    `=IFERROR(GOOGLEFINANCE(${c}; "high"); 0)`,                     // K: Day High
    `=IFERROR(GOOGLEFINANCE(${c}; "low"); 0)`,                      // L: Day Low
    `=IFERROR(GOOGLEFINANCE(${c}; "high52"); 0)`,                   // M: 52w High
    `=IFERROR(GOOGLEFINANCE(${c}; "low52"); 0)`,                    // N: 52w Low
    `=IFERROR(GOOGLEFINANCE(${c}; "beta"); "-")`                    // O: Beta
  ];

  // Write formulas in batch (horizontal range)
  // setFormulas expects a 2D array [[formula1, formula2, ...]]
  sheet.getRange(lastRow, 2, 1, 14).setFormulas([formulas]);

  // 3. Apply Formatting
  sheet.getRange(lastRow, 3).setNumberFormat("#,##0.00");     // Price
  sheet.getRange(lastRow, 4).setNumberFormat("0.00%");        // Percentage
  sheet.getRange(lastRow, 8).setNumberFormat("#,##0.00,, \"M\""); // Market Cap in Millions
  
  return "Analysis started: " + cleanTicker;
}

/**
 * Removes a row from the Watchlist based on its index.
 * * @param {number} row - The 1-based row index in the sheet.
 */
function removeFromWatchlist(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Watchlist");
  if (!sheet) throw new Error("Watchlist sheet not found");
  
  // Security Check: Only delete valid data rows (Row > 1)
  // Note: 'row' parameter comes from frontend (index + 2), matching actual sheet row.
  if (row && row > 1) {
    sheet.deleteRow(row);
  } else {
    throw new Error("Invalid row index");
  }
}