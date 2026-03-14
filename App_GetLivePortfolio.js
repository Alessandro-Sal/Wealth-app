/**
 * Retrieves live portfolio data including detailed metrics for Stocks, ETFs, and Crypto.
 * Reads extended data points from the "Stocks" sheet (Cols A-BD).
 * Reads extended data points from the "Crypto" sheet (Cols A-AK).
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
    // Read 56 columns (up to BD)
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 56).getDisplayValues();
    
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
        
        // --- TRADING DATA ---
        // AK (index 36) skipped
        firstBuyDate: r[37], // AL: First Buy Date
        firstPrice: r[38],   // AM: First Price
        minBuy: r[39],       // AN: Min Buy Price
        maxBuy: r[40],       // AO: Max Buy Price
        maxSell: r[41],      // AP: Max Sell Price
        avgSell: r[42],      // AQ: Avg Sell Price
        daysHeld: r[43],     // AR: Days Held
        lastActivity: r[44], // AS: Last Activity
        tradeCount: r[45],   // AT: Trade Count
        
        expectedDividend: r[46], // AU: Expected Dividend
        
        // --- NEW ADVANCED METRICS (Columns AV-BD) ---
        xirrUnrealized: r[47],     // AV: XIRR Unrealized
        xirrUnrealizedNote: r[48], // AW: Note Unrealized
        xirrRealized: r[49],       // AX: XIRR Realized
        xirrRealizedNote: r[50],   // AY: Note Realized
        twr: r[51],                // AZ: TWR (Modified Dietz)
        yoc: r[52],                // BA: Yield on Cost
        winRate: r[53],            // BB: Win Rate
        profitFactor: r[54],       // BC: Profit Factor
        drawdown: r[55]            // BD: Drawdown
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
      // Read up to column AK (Index 36) -> reading 37 cols total
      cryptoData = cSheet.getRange(3, 1, cSheet.getLastRow()-2, 37).getDisplayValues()
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
                
                // --- TRADING DATA (Columns T-AB) ---
                firstBuyDate: r[19],  // T: First Buy Date
                firstPrice: r[20],    // U: First Price
                minBuy: r[21],        // V: Min Buy Price
                maxBuy: r[22],        // W: Max Buy Price
                maxSell: r[23],       // X: Max Sell Price
                avgSell: r[24],       // Y: Avg Sell Price
                daysHeld: r[25],      // Z: Days Held
                lastActivity: r[26],  // AA: Last Activity
                tradeCount: r[27],    // AB: Trade Count
                
                // --- NEW ADVANCED METRICS (Columns AC-AK) ---
                xirrUnrealized: r[28],     // AC: XIRR Unrealized
                xirrUnrealizedNote: r[29], // AD: Note Unrealized
                xirrRealized: r[30],       // AE: XIRR Realized
                xirrRealizedNote: r[31],   // AF: Note Realized
                twr: r[32],                // AG: TWR (Modified Dietz)
                yoc: r[33],                // AH: Yield on Cost
                winRate: r[34],            // AI: Win Rate
                profitFactor: r[35],       // AJ: Profit Factor
                drawdown: r[36],           // AK: Drawdown
                
                // --- METADATA (Defaulting) ---
                div: "",              // No dividends column specified
                expectedDividend: "", 
                curr: "EUR",          
                name: r[0]            // Use Ticker as Name if not available
            };
        })
        .filter(i => i);
  }
// 3. Daily Portfolio Performance (from Dashboard)
  let dayChangeData = { val: "€ 0,00", pct: "0,00%", spyPct: 0 };
  const dashSheet = ss.getSheetByName("Stock Market Dashboard");
  if (dashSheet) {
    // Leggiamo la cella B8 per l'S&P 500 (Assicurati di inserire la formula in questa cella!)
    let spyValue = dashSheet.getRange("C5").getValue();
    
    // Se la cella è vuota o c'è un errore temporaneo di Google Finance, usiamo 0
    if (isNaN(spyValue) || spyValue === "") {
        spyValue = 0;
    }

    dayChangeData = { 
      val: dashSheet.getRange("B6").getDisplayValue(), 
      pct: dashSheet.getRange("B7").getDisplayValue(),
      spyPct: spyValue // <--- Il nuovo dato passato al Frontend
    };
  }

  return { stocks: stocks, etfs: etfs, crypto: cryptoData, dayChange: dayChangeData };
}