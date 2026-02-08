/**
 * Retrieves and structures analytics data from "Stocks" and "Crypto" sheets.
 * Maps specific columns for Value, Unrealized Gains, Realized Gains, and Dividends.
 * Automatically categorizes assets into Stocks, ETFs, or Crypto based on sheet origin and sector/ticker validation.
 * * @return {Object} An object containing arrays for { stocks, etfs, crypto }.
 */
function getAnalyticsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. STOCKS & ETFS (Sheet "Stocks")
  const sheetS = ss.getSheetByName("Stocks");
  let stocks = [];
  let etfs = [];

  if (sheetS && sheetS.getLastRow() >= 3) {
    // Read up to column W (23 columns total)
    // Mapping Indices: A=0, J=9, K=10, L=11, M=12, N=13, Q=16, U=20, V=21, W=22
    const dataS = sheetS.getRange(3, 1, sheetS.getLastRow() - 2, 23).getValues();
    
    dataS.forEach(r => {
      if (!r[0]) return; // Skip empty rows
      
      let item = {
        ticker: String(r[0]),
        value: parseFloat(r[9]) || 0,        // Col J (Current Value)
        gainPct: parseFloat(r[10]) || 0,     // Col K (Unrealized %)
        gainEur: parseFloat(r[11]) || 0,     // Col L (Unrealized €)
        
        // --- COLUMN CORRECTIONS ---
        // Realized Trading Only: NOW COLUMN N (Index 13)
        realizedExcl: parseFloat(r[13]) || 0, 
        
        // Dividends: Column Q (Index 16)
        dividends: parseFloat(r[16]) || 0,   

        equityPct: parseFloat(r[17]) || 0,   // Col R (Equity %)
        
        // --- CATEGORIES ---
        sector: r[20] ? String(r[20]).trim() : "Other",   // Col U
        industry: r[21] ? String(r[21]).trim() : "Other", // Col V (INDUSTRY ADDED)
        country: r[22] ? String(r[22]).trim() : "Global"  // Col W
      };

      // Filter Logic: Stocks vs ETFs
      // Checks explicit sector "ETFs" or specific ticker keywords
      if (item.sector === "ETFs" || item.ticker.includes("EMA") || item.ticker.includes("VWCE")) { 
        etfs.push(item);
      } else {
        stocks.push(item);
      }
    });
  }

  // 2. CRYPTO (Sheet "Crypto")
  const sheetC = ss.getSheetByName("Crypto");
  let crypto = [];
  if (sheetC && sheetC.getLastRow() >= 3) {
    const dataC = sheetC.getRange(3, 1, sheetC.getLastRow() - 2, 13).getValues();
    dataC.forEach(r => {
      if (!r[0]) return;
      crypto.push({
        ticker: String(r[0]),
        value: parseFloat(r[6]) || 0,
        gainPct: parseFloat(r[7]) || 0,
        gainEur: parseFloat(r[8]) || 0,
        realizedExcl: 0, 
        dividends: 0,
        equityPct: parseFloat(r[9]) || 0,
        sector: r[12] ? String(r[12]).trim() : "Crypto",
        industry: "Digital Assets",
        country: "Global"
      });
    });
  }

  return { stocks: stocks, etfs: etfs, crypto: crypto };
}