/**
 * Fetches closed positions data from 'Graveyard_Stocks' and 'Graveyard_Crypto'.
 * Returns a unified array of objects representing each closed position.
 */
function getClosedPositionsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stocksSheet = ss.getSheetByName("Graveyard_Stocks");
    const cryptoSheet = ss.getSheetByName("Graveyard_Crypto");

    const result = [];

    // --- PROCESS STOCKS & ETFS (Columns A to AS) ---
    if (stocksSheet) {
      const data = stocksSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0] || row[0] === "Tickers" || row[0] === "") continue;

        let type = 'Stock';
        const sec = String(row[13] || '').toUpperCase();
        const ind = String(row[14] || '').toUpperCase();
        if (sec.includes('ETF') || ind.includes('ETF') || sec.includes('FUND')) {
          type = 'ETF';
        }

        result.push({
          ticker: row[0] || "-",          // A
          type: type,
          qty: row[1] || "0",             // B: Shares
          avgSell: row[2] || row[35] || "-", // C: Avg Sell Price
          priceToday: row[4] || "-",      // E: Price (Euro)
          missedPct: row[5] || "0,00%",   // F: Change% (Opportunity Cost)
          missedEur: row[6] || "€ 0,00",  // G: Change£ (Opportunity Cost)
          invested: row[7] || "€ 0,00",   // H
          realizedPct: row[8] || "0,00%", // I
          realPnL: row[9] || "€ 0,00",    // J
          divs: row[10] || "€ 0,00",      // K
          roi: row[11] || "0,00%",        // L
          totalPnL: row[12] || "€ 0,00",  // M
          sector: row[13] || "-",         // N
          ind: row[14] || "-",            // O
          ctry: row[15] || "-",           // P
          name: row[17] || "-",           // R
          curr: row[18] || "-",           // S
          pe: row[19] || "-",             // T
          eps: row[20] || "-",            // U
          mkt: row[21] || "-",            // V
          beta: row[28] || "-",           // AC
          firstBuy: row[30] || "-",       // AE
          firstPrice: row[31] || "-",     // AF
          minBuy: row[32] || "-",         // AG
          maxBuy: row[33] || "-",         // AH
          maxSell: row[34] || "-",        // AI
          daysHeld: row[36] || "0",       // AK
          lastAct: row[37] || "-",        // AL
          trades: row[38] || "0",         // AM
          xirr: row[39] || "0,00%",       // AN
          xirrNote: row[40] || "-",       // AO
          twr: row[41] || "0,00%",        // AP
          yoc: row[42] || "-",            // AQ
          winRate: row[43] || "-",        // AR
          profitFactor: row[44] || "-"    // AS
        });
      }
    }

    // --- PROCESS CRYPTO (Columns A to V) ---
    if (cryptoSheet) {
      const data = cryptoSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0] || row[0] === "Tickers" || row[0] === "") continue;

        result.push({
          ticker: row[0] || "-",          
          type: 'Crypto',
          qty: row[1] || "0",             // B: Shares
          avgSell: row[2] || row[15] || "-", // C: Avg Sell Price
          priceToday: row[4] || "-",      // E: Price (Euro)
          missedPct: row[5] || "0,00%",   // F: Change% 
          missedEur: row[6] || "€ 0,00",  // G: Change£
          invested: row[7] || "€ 0,00",   // H
          realizedPct: row[8] || "0,00%", // I: Total Gain%
          realPnL: row[9] || "€ 0,00",    // J: Total Gain
          divs: "€ 0,00",                 
          roi: row[8] || "0,00%",         
          totalPnL: row[9] || "€ 0,00",   
          sector: "-", ind: "-", ctry: "-", name: "-", curr: "-", pe: "-", eps: "-", mkt: "-", beta: "-",                      
          firstBuy: row[10] || "-",       // K
          firstPrice: row[11] || "-",     // L
          minBuy: row[12] || "-",         // M
          maxBuy: row[13] || "-",         // N
          maxSell: row[14] || "-",        // O
          daysHeld: row[16] || "0",       // Q
          lastAct: row[17] || "-",        // R
          trades: row[18] || "0",         // S
          xirr: row[19] || "0,00%",       // T
          xirrNote: row[20] || "-",       // U
          twr: row[21] || "0,00%",        // V
          yoc: "-", winRate: "-", profitFactor: "-"
        });
      }
    }

    return result;
  } catch (e) {
    return { error: e.message };
  }
}