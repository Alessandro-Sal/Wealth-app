/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Computes a true mathGrandTotal from scratch to guarantee 100% distribution allocation.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  SpreadsheetApp.flush(); 

  const isCalculating = (rawVal, displayVal) => {
    const str = String(displayVal).toUpperCase();
    return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
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

  if (isCalculating(cryptoRaw, cryptoStr)) {
    return { error: "Sheet is recalculating. Skipping update." };
  }

  // --- 1. ESTRATTORE NUMERICO BLINDATO ---
  const getSafeNum = (row, col) => {
    let val = sheet.getRange(row, col).getValue();
    if (typeof val === 'number') return val;
    if (!val) return 0;
    return parseFloat(String(val).replace(/[^0-9,-]+/g,"").replace(',', '.')) || 0;
  };

  // --- 2. LETTURA COMPONENTI INDIVIDUALI ---
  const valEtfs = getSafeNum(2, 2);
  const valStocks = getSafeNum(3, 2);
  const valCash = getSafeNum(4, 2);
  const valCashEq = getSafeNum(5, 2);
  const valCrypto = getSafeNum(6, 2);
  const valOthers = getSafeNum(7, 2);
  const valPension = getSafeNum(24, 4);

  let realAssets;
  try { realAssets = getRealAssetsSummary(); } catch(e) {}
  if (!realAssets) realAssets = { realEstate: { net: 0 }, bonds: { net: 0 }, totalNetWorthImpact: 0 };

  // --- 3. COSTRUZIONE VERO TOTALE MATEMATICO ---
  // Per il totale usiamo CashEq se esiste, altrimenti il Cash normale (evita doppi conteggi)
  const effectiveCash = valCashEq > 0 ? valCashEq : valCash;
  const mathLiquidNW = valStocks + valEtfs + effectiveCash + valCrypto + valOthers;
  const mathGrandTotal = mathLiquidNW + valPension + realAssets.totalNetWorthImpact;

  const fmt = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });

  // --- 4. CALCOLO PERCENTUALI AL 100% ---
  const calcPct = (raw) => mathGrandTotal > 0 ? ((raw / mathGrandTotal) * 100).toFixed(1) + "%" : "0.0%";

  const getRecalcRow = (row, rawNum) => {
    return { amount: sheet.getRange(row, 2).getDisplayValue(), raw: rawNum, percent: calcPct(rawNum) };
  };

  const getSectionData = (startRow) => {
    return {
      unrealized: { amount: sheet.getRange(startRow, 2).getDisplayValue(), percent: sheet.getRange(startRow, 3).getDisplayValue() },
      realized: { amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), percent: sheet.getRange(startRow + 1, 3).getDisplayValue() },
      balance: { amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), percent: sheet.getRange(startRow + 2, 3).getDisplayValue() },
      invested: { amount: sheet.getRange(startRow + 3, 2).getDisplayValue(), percent: "" }
    };
  };

  return {
    liquidNetWorth: fmt.format(mathLiquidNW),    
    liquidNetWorthUSD: sheet.getRange(26, 2).getDisplayValue(), 
    totalNetWorth: fmt.format(mathGrandTotal), // USA IL VERO TOTALE
    totalNetWorthUSD: sheet.getRange(24, 2).getDisplayValue(), 

    summary: { 
      grandTotal: mathGrandTotal, 
      etfs: getRecalcRow(2, valEtfs),      
      stocks: getRecalcRow(3, valStocks),    
      cash: getRecalcRow(4, valCash),      
      cashEq: getRecalcRow(5, valCashEq),    
      crypto: getRecalcRow(6, valCrypto),    
      others: getRecalcRow(7, valOthers),
      pension: { amount: sheet.getRange(24, 4).getDisplayValue(), raw: valPension, percent: calcPct(valPension) },
      realEstate: { amount: fmt.format(realAssets.realEstate.net), raw: realAssets.realEstate.net, percent: calcPct(realAssets.realEstate.net) },
      bonds: { amount: fmt.format(realAssets.bonds.net), raw: realAssets.bonds.net, percent: calcPct(realAssets.bonds.net) }
    },

    cryptoSection: { main: getRecalcRow(6, valCrypto), ...getSectionData(9) },
    stocksSection: { main: getRecalcRow(3, valStocks), ...getSectionData(14) },
    etfSection: { main: getRecalcRow(2, valEtfs), ...getSectionData(19) }
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