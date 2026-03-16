/**
 * Recupera i dati di evoluzione del portafoglio dal foglio "Storico" per un anno specifico.
 * @param {string|number} year L'anno da filtrare (es. "2026")
 */
function getHistoricalPortfolio(year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Storico");
    
    if (!sheet) return { error: "Sheet 'Storico' not found" };

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; // Foglio vuoto oltre le intestazioni

    const targetYear = String(year).trim();
    
    // Filtro con conversione sicura a stringa
    const filtered = data.slice(1).filter(row => {
      const rowYear = String(row[0]).trim();
      const ticker = String(row[2]).trim();
      return rowYear === targetYear && ticker !== "";
    });

    return filtered.map(row => ({
      year: row[0],
      date: row[1],
      ticker: row[2],
      type: row[3],
      qty: row[4],
      invested: row[5],
      priceOrig: row[6],
      priceEur: row[7],
      value: row[8],
      gainLoss: row[9],
      performance: String(row[10] || "0%") // Previene undefined
    }));
    
  } catch (e) {
    return { error: e.toString() };
  }
}