/**
 * Fetches data to populate autocomplete suggestions in the UI.
 * 1. Builds a map of Categories -> Historical Notes (to suggest descriptions based on category).
 * 2. Compiles a list of unique Asset Tickers (from a default list + actual history).
 * * @return {Object} Object containing { notes: {Category: [Notes...]}, tickers: [BTC, AAPL...] }
 */
function getAutocompleteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. EXPENSE NOTES AUTOCOMPLETE ---
  // Currently points to 2026. Ideally, this could scan the current year dynamically.
  const expSheet = ss.getSheetByName("Expenses Tracker 2026");
  let notesMap = {};

  if (expSheet && expSheet.getLastRow() >= 20) {
    // Read Category (Col C) and Note (Col D) starting from Row 20
    const data = expSheet.getRange(20, 3, expSheet.getLastRow() - 19, 2).getValues();
    
    data.forEach(r => {
      const category = r[0];
      const note = r[1];
      
      if (category && note) {
        if (!notesMap[category]) {
            notesMap[category] = [];
        }
        // Avoid duplicates
        if (!notesMap[category].includes(note)) {
            notesMap[category].push(note);
        }
      }
    });
  }

  // --- 2. TICKER AUTOCOMPLETE ---
  // Start with a set of popular default tickers
  let tickers = new Set(["BTC", "ETH", "VWCE", "AAPL", "NVDA", "MSFT", "TSLA", "AMZN", "GOOGL", "META", "S&P500"]);

  // Helper function to scan History sheets for existing assets
  const scanSheet = (name) => {
    const s = ss.getSheetByName(name);
    if(s && s.getLastRow() >= 3) {
      // Read Ticker Column (Column B, starting Row 3)
      const vals = s.getRange(3, 2, s.getLastRow() - 2, 1).getValues();
      vals.forEach(v => {
        if(v[0]) tickers.add(String(v[0]).toUpperCase().trim());
      });
    }
  };

  // Scan both Stock and Crypto history
  scanSheet("History B/S Stocks"); 
  scanSheet("History B/S Crypto");

  return { 
    notes: notesMap, 
    tickers: Array.from(tickers).sort() 
  };
}

/**
 * Closes all active UI modal sheets and the backdrop.
 * Centralized function to ensure no overlays remain open.
 */
function closeAllSheets() {
    // Complete list of all panels/modals to close
    const sheets = [
        'calc-sheet', 
        'search-sheet', 
        'invest-search-sheet', 
        'subs-sheet', 
        'gemini-sheet', // AI Assistant
        'risk-sheet'    // Risk Analysis Report
    ];

    // Iterate through all and remove the 'active' class
    sheets.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });

    // Close the dark backdrop overlay
    const backdrop = document.getElementById('common-backdrop');
    if (backdrop) backdrop.classList.remove('active');
}

/**
 * Ottiene il tasso di cambio e converte un importo.
 * Sfrutta GoogleFinance tramite una formula temporanea sul foglio o una chiamata API.
 * @param {number} amount L'importo da convertire.
 * @param {string} fromCurrency La valuta di origine (es. "USD").
 * @param {string} toCurrency La valuta di destinazione (es. "EUR").
 * @returns {number|string} L'importo convertito o una stringa di errore.
 */
function getConversionRate(amount, fromCurrency, toCurrency) {
  if (!amount || !fromCurrency || !toCurrency) return "Err: Parametri mancanti";
  if (fromCurrency === toCurrency) return amount;

  try {
    // URL per l'API pubblica (ad es. un servizio gratuito o Google Finance se accessibile via UrlFetchApp)
    // Usiamo qui un approccio alternativo veloce tramite un API aperta per i tassi di cambio
    const url = `https://api.exchangerate-api.com/v4/latest/${fromCurrency}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) {
      return "Err: Impossibile recuperare il tasso";
    }
    
    const data = JSON.parse(response.getContentText());
    if (data && data.rates && data.rates[toCurrency]) {
      const rate = data.rates[toCurrency];
      const convertedAmount = amount * rate;
      return convertedAmount.toFixed(2);
    } else {
      return "Err: Tasso non trovato";
    }
  } catch (e) {
    return "Err: " + e.message;
  }
}