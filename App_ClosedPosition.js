/**
 * Fetches closed positions data from 'DashboardAnalitica' sheet.
 * Returns an array of objects representing each closed position.
 */
function getClosedPositionsData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DashboardAnalitica");
    if (!sheet) throw new Error("Sheet 'DashboardAnalitica' not found");

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const result = [];
    
    // Start from 1 to skip headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Filter only CLOSED positions (Column C is index 2)
      if (row[2] === "CLOSED") {
        result.push({
          ticker: row[0] || "-",          // A
          type: row[1] || "-",            // B
          invested: row[5] || "€ 0,00",   // F
          realPnL: row[6] || "€ 0,00",    // G
          divs: row[7] || "€ 0,00",       // H
          totalPnL: row[8] || "€ 0,00",   // I
          roi: row[13] || "0,00%",        // N
          firstBuy: row[14] || "-",       // O
          firstPrice: row[15] || "-",     // P
          minBuy: row[16] || "-",         // Q
          maxBuy: row[17] || "-",         // R
          maxSell: row[18] || "-",        // S
          avgSell: row[19] || "-",        // T
          daysHeld: row[20] || "0",       // U
          lastAct: row[21] || "-",        // V
          trades: row[22] || "0",         // W
          xirr: row[26] || "0,00%",       // AA
          xirrNote: row[27] || "-",       // AB
          twr: row[28] || "0,00%"         // AC
        });
      }
    }
    
    return result;
  } catch (e) {
    return { error: e.message };
  }
}