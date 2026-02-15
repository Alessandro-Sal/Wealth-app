/**
 * Retrieves live portfolio data including detailed metrics for Stocks, ETFs, and Crypto.
 * Reads extended data points from the "Stocks" sheet (Cols A-AT).
 * Reads extended data points from the "Crypto" sheet (Cols A-S).
 * Fetches daily global performance from the Dashboard.
 * @return {Object} Structured data object: { stocks, etfs, crypto, dayChange }.
 */
function getLivePortfolio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. STOCKS & ETFS
  const sheet = ss.getSheetByName("Stocks");
  let stocks = [];
  let etfs = [];
  
  if (sheet && sheet.getLastRow() >= 3) {
    // MODIFICA: Leggiamo ora 46 colonne (da A ad AT) invece di 40
    // A=1 ... AJ=36 ... AT=46
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 46).getDisplayValues();
    
    const allItems = data.map(r => {
      // Skip empty rows or assets with 0 value
      if (!r[0] || r[9] === "0" || r[9] === "") return null;
      
      return {
        t: r[0],          // A: Ticker
        qty: r[1],        // B: Quantity
        avgPrice: r[2],   // C: Average Price
        price: r[3],      // D: Current Price
        pct: r[5],        // F: Change %
        dCh: r[7],        // H: Day Change
        val: r[9],        // J: Value
        
        // --- PERFORMANCE METRICS ---
        unrealizedPct: r[10], // K: Unrealized Gain %
        unrealizedEur: r[11], // L: Unrealized Gain €
        realizedPct: r[12],   // M: Realized Gain %
        realizedEur: r[13],   // N: Realized Gain €
        balPct: r[14],        // O: Balance %
        balEur: r[15],        // P: Balance €
        div: r[16],           // Q: Dividends
        
        bep: r[18],           // S: BreakEven Point
        delta: r[19],         // T: Delta
        
        // --- METADATA & EXTENDED DATA ---
        sector: r[20],    // U: Sector
        ind: r[21],       // V: Industry
        ctry: r[22],      // W: Country
        ex: r[23],        // X: Exchange
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
        beta: r[35],      // AJ: Beta
        
        // --- NEW TRADING DATA (Columns AL-AT) ---
        // AK (index 36) skipped/unused based on your request
        firstBuyDate: r[37], // AL: First Buy Date
        firstPrice: r[38],   // AM: First Price
        minBuy: r[39],       // AN: Min Buy Price
        maxBuy: r[40],       // AO: Max Buy Price
        maxSell: r[41],      // AP: Max Sell Price
        avgSell: r[42],      // AQ: Avg Sell Price
        daysHeld: r[43],     // AR: Days Held
        lastActivity: r[44], // AS: Last Activity
        tradeCount: r[45]    // AT: Trade Count
      };
    }).filter(i => i !== null);

    // Separate Stocks from ETFs based on Sector column
    stocks = allItems.filter(i => i.sector !== 'ETFs');
    etfs = allItems.filter(i => i.sector === 'ETFs');
  }

  // 2. CRYPTO
  let cryptoData = [];
  const cSheet = ss.getSheetByName("Crypto");
  if(cSheet && cSheet.getLastRow() >= 3) {
      // Read up to column S (Index 18) -> reading 20 cols to be safe
      cryptoData = cSheet.getRange(3, 1, cSheet.getLastRow()-2, 20).getDisplayValues()
        .map(r => {
            if(!r[0]) return null; // Skip if no ticker
            
            return { 
                t: r[0],              // A: Ticker
                qty: r[1],            // B: Quantity
                avgPrice: r[2],       // C: Average Price
                price: r[3],          // D: Current Price
                pct: null,            // No Day Change % for Crypto
                val: r[6],            // G: Current Value
                
                // --- CRYPTO SPECIFIC MAPPING ---
                unrealizedPct: r[7],  // H: Unrealized Gain %
                unrealizedEur: r[8],  // I: Unrealized Gain €
                
                realizedPct: r[13],   // N: Realized Gain %
                realizedEur: r[14],   // O: Realized Gain €
                
                balPct: r[15],        // P: Total Gain % (Mapped to Balance %)
                balEur: r[16],        // Q: Total Gain € (Mapped to Balance €)
                
                bep: r[17],           // R: BreakEven Point
                delta: r[18],         // S: Delta
                
                // --- METADATA (Defaulting) ---
                div: "",              // No dividends column specified
                curr: "EUR",          
                name: r[0]            // Use Ticker as Name if not available
            };
        })
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