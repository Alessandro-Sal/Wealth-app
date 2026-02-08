/**
 * Retrieves live portfolio data including detailed metrics for Stocks, ETFs, and Crypto.
 * Reads extended data points (P/E, Beta, 52w High/Low) from the "Stocks" sheet (Cols A-AJ).
 * Fetches daily global performance from the Dashboard.
 * * @return {Object} Structured data object: { stocks, etfs, crypto, dayChange }.
 */
function getLivePortfolio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. STOCKS & ETFS (Extended read up to column 36 = AJ)
  const sheet = ss.getSheetByName("Stocks");
  let stocks = [];
  let etfs = [];
  
  if (sheet && sheet.getLastRow() >= 3) {
    // Read raw data starting from Row 3, Column 1 (A) for 36 columns (up to AJ)
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 36).getDisplayValues();
    
    const allItems = data.map(r => {
      // Skip empty rows or assets with 0 value
      if (!r[0] || r[9] === "0" || r[9] === "") return null;
      
      return {
        t: r[0],          // A: Ticker
        price: r[3],      // D: Price
        val: r[9],        // J: Value
        pct: r[5],        // F: Change %
        dCh: r[7],        // H: Day Change
        sector: r[20],    // U: Sector
        ex: r[23],        // X: Exchange
        
        // --- EXTENDED DATA (U -> AJ) ---
        ind: r[21],       // V: Industry
        ctry: r[22],      // W: Country
        name: r[24],      // Y: Name
        curr: r[25],      // Z: Currency
        pe: r[26],        // AA: P/E
        eps: r[27],       // AB: EPS
        mkt: r[28],       // AC: Market Cap
        vol: r[29],       // AD: Volume
        avgVol: r[30],    // AE: Avg Vol
        dayH: r[31],      // AF: Day High
        dayL: r[32],      // AG: Day Low
        yearH: r[33],     // AH: 52w High
        yearL: r[34],     // AI: 52w Low
        beta: r[35]       // AJ: Beta
      };
    }).filter(i => i !== null);

    // Separate Stocks from ETFs based on Sector column
    stocks = allItems.filter(i => i.sector !== 'ETFs');
    etfs = allItems.filter(i => i.sector === 'ETFs');
  }

  // 2. CRYPTO (Standard read)
  let cryptoData = [];
  const cSheet = ss.getSheetByName("Crypto");
  if(cSheet && cSheet.getLastRow() >= 3) {
      cryptoData = cSheet.getRange(3, 1, cSheet.getLastRow()-2, 7).getDisplayValues()
        .map(r => r[0] ? { t: r[0], price: r[3], val: r[6], pct: null } : null)
        .filter(i => i);
  }

  // 3. Daily Portfolio Performance (from Dashboard)
  let dayChangeData = { val: "€ 0,00", pct: "0,00%" };
  const dashSheet = ss.getSheetByName("Stock Market Dashboard");
  if (dashSheet) {
    dayChangeData = { 
      val: dashSheet.getRange("B6").getDisplayValue(), 
      pct: dashSheet.getRange("B7").getDisplayValue() 
    };
  }

  return { stocks: stocks, etfs: etfs, crypto: cryptoData, dayChange: dayChangeData };
}