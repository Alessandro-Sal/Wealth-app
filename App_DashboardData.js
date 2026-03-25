/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Includes Real Assets & Liabilities to calculate a global holistic Net Worth and accurate allocation percentages.
 * @return {Object} Dashboard data object containing summary stats and section-specific details.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  SpreadsheetApp.flush(); 

  const isCalculating = (rawVal, displayVal) => {
    const str = String(displayVal).toUpperCase();
    const hasErrorString = str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
    const isTemporaryZero = (rawVal === 0 || rawVal === "0" || str === "€ 0,00" || str === "0,00 €");
    return hasErrorString || isTemporaryZero;
  };

  let cryptoRaw = sheet.getRange(6, 2).getValue();
  let cryptoStr = sheet.getRange(6, 2).getDisplayValue();
  let retries = 0;
  
  while (isCalculating(cryptoRaw, cryptoStr) && retries < 1) {
    Utilities.sleep(1000); 
    SpreadsheetApp.flush(); 
    cryptoRaw = sheet.getRange(6, 2).getValue();
    cryptoStr = sheet.getRange(6, 2).getDisplayValue();
    retries++;
  }

  const isStillError = (displayVal) => {
     const str = String(displayVal).toUpperCase();
     return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
  };

  if (isStillError(cryptoStr)) {
    return { error: "Sheet is recalculating or API down. Skipping update." };
  }

  // --- 1. INTEGRAZIONE ASSET REALI ALLA RADICE ---
  let realAssets = getRealAssetsSummary() || { 
      realEstate: { net: 0 }, 
      bonds: { net: 0 }, 
      totalNetWorthImpact: 0 
  };

  // Calcolo VERO GRAND TOTAL (Liquidità + Immobili + Bond - Debiti)
  let rawBaseTotal = parseFloat(sheet.getRange(24, 1).getValue()) || 0;
  let grandTotal = rawBaseTotal + realAssets.totalNetWorthImpact;
  
  const fmt = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });

  // --- 2. RICALCOLO PERCENTUALI SUL NUOVO TOTALONE ---
  const getRecalculatedRow = (row) => {
    const val = sheet.getRange(row, 2).getDisplayValue(); 
    const raw = parseFloat(sheet.getRange(row, 2).getValue()) || 0;
    let pct = grandTotal > 0 ? ((raw / grandTotal) * 100).toFixed(1) + "%" : "0.0%";
    return { amount: val, raw: raw, percent: pct };
  };

  const getSectionData = (startRow) => {
    return {
      unrealized: { amount: sheet.getRange(startRow, 2).getDisplayValue(), percent: sheet.getRange(startRow, 3).getDisplayValue() },
      realized: { amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), percent: sheet.getRange(startRow + 1, 3).getDisplayValue() },
      balance: { amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), percent: sheet.getRange(startRow + 2, 3).getDisplayValue() },
      invested: { amount: sheet.getRange(startRow + 3, 2).getDisplayValue(), percent: "" }
    };
  };

  const pensionRaw = parseFloat(sheet.getRange(24, 4).getValue()) || 0;

  return {
    liquidNetWorth: sheet.getRange(26, 1).getDisplayValue(),    
    liquidNetWorthUSD: sheet.getRange(26, 2).getDisplayValue(), 
    
    totalNetWorth: fmt.format(grandTotal), // <-- RESTITUISCE IL NUMERONE AGGIORNATO AL 100%
    totalNetWorthUSD: sheet.getRange(24, 2).getDisplayValue(), 

    summary: { 
      etfs: getRecalculatedRow(2),      
      stocks: getRecalculatedRow(3),    
      cash: getRecalculatedRow(4),      
      cashEq: getRecalculatedRow(5),    
      crypto: getRecalculatedRow(6),    
      others: getRecalculatedRow(7),
      pension: {
        amount: sheet.getRange(24, 4).getDisplayValue(), 
        raw: pensionRaw,
        percent: grandTotal > 0 ? ((pensionRaw / grandTotal) * 100).toFixed(1) + "%" : "0.0%"
      },
      // Inseriamo anche Real Estate e Bonds nel summary con la % calcolata!
      realEstate: {
        amount: fmt.format(realAssets.realEstate.net),
        raw: realAssets.realEstate.net,
        percent: grandTotal > 0 ? ((realAssets.realEstate.net / grandTotal) * 100).toFixed(1) + "%" : "0.0%"
      },
      bonds: {
        amount: fmt.format(realAssets.bonds.net),
        raw: realAssets.bonds.net,
        percent: grandTotal > 0 ? ((realAssets.bonds.net / grandTotal) * 100).toFixed(1) + "%" : "0.0%"
      }
    },

    cryptoSection: { main: getRecalculatedRow(6), ...getSectionData(9) },
    stocksSection: { main: getRecalculatedRow(3), ...getSectionData(14) },
    etfSection: { main: getRecalculatedRow(2), ...getSectionData(19) }
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