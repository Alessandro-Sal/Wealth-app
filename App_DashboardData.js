/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Maps specific rows for Total and Liquid Net Worth based on user configuration.
 * * @return {Object} Dashboard data object containing summary stats and section-specific details.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  // --- ULTIMATE FIX: FORCE CALCULATION & CATCH "0" ---
  // 1. Force Google Sheets to apply any pending formulas immediately
  SpreadsheetApp.flush(); 

  // 2. Check if it's calculating (Errors, "Loading", or even temporarily "0")
  const isCalculating = (rawVal, displayVal) => {
    const str = String(displayVal).toUpperCase();
    const hasErrorString = str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
    
    // Many crypto formulas (like IFERROR) return 0 while loading from CoinGecko/Binance.
    // We treat 0 as a "temporary" state and wait a few seconds just to be sure.
    const isTemporaryZero = (rawVal === 0 || rawVal === "0" || str === "€ 0,00" || str === "0,00 €");
    
    return hasErrorString || isTemporaryZero;
  };

  let cryptoRaw = sheet.getRange(6, 2).getValue();
  let cryptoStr = sheet.getRange(6, 2).getDisplayValue();
  let retries = 0;
  
  // OPTIMIZED: Reduced wait time to prevent blocking the UI.
  // We wait max 1 second instead of 6. The 60-second polling will catch any late API responses.
  while (isCalculating(cryptoRaw, cryptoStr) && retries < 1) {
    Utilities.sleep(1000); 
    SpreadsheetApp.flush(); // Force refresh at each tick to get fresh API data
    cryptoRaw = sheet.getRange(6, 2).getValue();
    cryptoStr = sheet.getRange(6, 2).getDisplayValue();
    retries++;
  }

  // If it's still showing an ERROR string after 6s, abort to protect UI.
  // (Note: If it's still '0' after 6 seconds, we accept it, because you might genuinely have 0 balance)
  const isStillError = (displayVal) => {
     const str = String(displayVal).toUpperCase();
     return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
  };

  if (isStillError(cryptoStr)) {
    console.log("Crypto values are in error state. Skipping update.");
    return { error: "Sheet is recalculating or API down. Skipping update." };
  }

  // Helper to read Value (Col B=2) and Percentage (Col C=3) for top tables
  const getRowData = (row) => {
    const val = sheet.getRange(row, 2).getDisplayValue(); 
    const raw = sheet.getRange(row, 2).getValue();
    const pct = sheet.getRange(row, 3).getDisplayValue(); 
    return { amount: val, raw: (typeof raw === 'number' ? raw : 0), percent: pct };
  };

  const getSectionData = (startRow) => {
    return {
      unrealized: { 
        amount: sheet.getRange(startRow, 2).getDisplayValue(),     
        percent: sheet.getRange(startRow, 3).getDisplayValue() 
      },
      realized: { 
        amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), 
        percent: sheet.getRange(startRow + 1, 3).getDisplayValue() 
      },
      balance: { 
        amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), 
        percent: sheet.getRange(startRow + 2, 3).getDisplayValue() 
      },
      invested: {
        amount: sheet.getRange(startRow + 3, 2).getDisplayValue(), 
        percent: "" 
      }
    };
  };

  return {
    // Liquid NW at Row 26 (Col A=1 EUR, Col B=2 USD)
    liquidNetWorth: sheet.getRange(26, 1).getDisplayValue(),    
    liquidNetWorthUSD: sheet.getRange(26, 2).getDisplayValue(), 
    
    // Total NW at Row 24 (Col A=1 EUR, Col B=2 USD)
    totalNetWorth: sheet.getRange(24, 1).getDisplayValue(),     
    totalNetWorthUSD: sheet.getRange(24, 2).getDisplayValue(),  

    summary: { 
      etfs: getRowData(2),      
      stocks: getRowData(3),    
      cash: getRowData(4),      
      cashEq: getRowData(5),    
      crypto: getRowData(6),    
      others: getRowData(7),
      pension: {
        amount: sheet.getRange(24, 4).getDisplayValue(), 
        raw: (typeof sheet.getRange(24, 4).getValue() === 'number' ? sheet.getRange(24, 4).getValue() : 0),
        percent: sheet.getRange(24, 5).getDisplayValue() 
      }
    },

    cryptoSection: {
      main: getRowData(6), 
      ...getSectionData(9) 
    },
    stocksSection: {
      main: getRowData(3),
      ...getSectionData(14)
    },
    etfSection: {
      main: getRowData(2),
      ...getSectionData(19)
    }
  };
}

/**
 * Recupera la sintesi degli Asset Reali e delle Passività dall'ultimo mese registrato nel Log.
 * Da richiamare per popolare le card della Dashboard.
 */
function getRealAssetsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log_Valuations");
  if (!logSheet) return null;

  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  // 1. Trova l'ultimo mese registrato
  let latestDate = new Date(0);
  for (let i = 1; i < data.length; i++) {
     let d = new Date(data[i][0]);
     if (!isNaN(d.getTime()) && d > latestDate) latestDate = d;
  }
  let latestMonthStr = Utilities.formatDate(latestDate, Session.getScriptTimeZone(), "yyyy-MM");

  // 2. Prepara l'oggetto di sintesi
  let summary = {
     realEstate: { gross: 0, debt: 0, net: 0 },
     bonds: { gross: 0, debt: 0, net: 0 },
     liabilities: { gross: 0, debt: 0, net: 0 }, // Debiti puri (senza asset collegato)
     totalNetWorthImpact: 0
  };

  // 3. Somma i dati dell'ultimo mese
  for (let i = 1; i < data.length; i++) {
     let rowDate = new Date(data[i][0]);
     if (isNaN(rowDate.getTime())) continue;
     
     if (Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM") === latestMonthStr) {
        let type = data[i][1];
        let gross = parseFloat(data[i][3]) || 0;
        let debt = parseFloat(data[i][4]) || 0;
        let net = parseFloat(data[i][5]) || 0;

        if (type === "Real Estate") {
           summary.realEstate.gross += gross;
           summary.realEstate.debt += debt;
           summary.realEstate.net += net;
        } else if (type.includes("Bond")) {
           summary.bonds.gross += gross;
           summary.bonds.debt += debt;
           summary.bonds.net += net;
        } else if (type === "Liability") {
           summary.liabilities.gross += gross;
           summary.liabilities.debt += debt;
           summary.liabilities.net += net;
        }
        
        // Calcola l'impatto totale sul patrimonio netto
        summary.totalNetWorthImpact += net;
     }
  }
  
  return summary;
}