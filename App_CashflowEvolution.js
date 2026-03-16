/**
 * Fetches cashflow data from the 'Portfolio_Flussi' sheet.
 * Returns an array of objects representing each row.
 */
function getPortfolioCashflowData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Portfolio_Flussi");
    if (!sheet) throw new Error("Sheet 'Portfolio_Flussi' not found");

    // Get display values to preserve currency formatting directly from the sheet
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const result = [];
    
    // Start from 1 to skip headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows

      result.push({
        year: row[0],
        assetType: row[1],
        deposits: row[2],
        withdrawals: row[3],
        netDividends: row[4],
        totalInflow: row[5],
        netFlow: row[6]
      });
    }
    
    return result;
  } catch (e) {
    return { error: e.message };
  }
}