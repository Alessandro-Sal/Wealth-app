/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Implements strict definitions for Liquid vs Total NW and a Cache Fallback for Crypto APIs.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  SpreadsheetApp.flush(); 

  const isCalculating = (rawVal, displayVal) => {
    const str = String(displayVal).toUpperCase();
    return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A") || rawVal === 0 || str === "€ 0,00";
  };

  let cryptoRaw = sheet.getRange(6, 2).getValue();
  let cryptoStr = sheet.getRange(6, 2).getDisplayValue();
  let retries = 0;
  
  // Aumentati i retry a 4 secondi per dare tempo alle API esterne
  while (isCalculating(cryptoRaw, cryptoStr) && retries < 4) {
    Utilities.sleep(1000); 
    SpreadsheetApp.flush(); 
    cryptoRaw = sheet.getRange(6, 2).getValue();
    cryptoStr = sheet.getRange(6, 2).getDisplayValue();
    retries++;
  }

  // --- FIX PUNTO 4: CACHE DI EMERGENZA CRYPTO ---
  const cache = CacheService.getUserCache();
  if (!isCalculating(cryptoRaw, cryptoStr) && cryptoRaw > 0) {
      // Se il dato è buono, lo salviamo in memoria per 6 ore
      cache.put('last_crypto_raw', cryptoRaw.toString(), 21600);
      cache.put('last_crypto_str', cryptoStr, 21600);
  } else {
      // Se le API sono ancora rotte/a zero, peschiamo l'ultimo dato noto!
      const cachedRaw = cache.get('last_crypto_raw');
      const cachedStr = cache.get('last_crypto_str');
      if (cachedRaw) {
          cryptoRaw = parseFloat(cachedRaw);
          cryptoStr = cachedStr;
      }
  }

  const getSafeNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    return parseFloat(String(val).replace(/[^0-9,-]+/g,"").replace(',', '.')) || 0;
  };

  const valEtfs = getSafeNum(sheet.getRange(2, 2).getValue());
  const valStocks = getSafeNum(sheet.getRange(3, 2).getValue());
  const valCash = getSafeNum(sheet.getRange(4, 2).getValue());
  const valCashEq = getSafeNum(sheet.getRange(5, 2).getValue());
  const valCrypto = getSafeNum(cryptoRaw);
  const valOthers = getSafeNum(sheet.getRange(7, 2).getValue());
  const valPension = getSafeNum(sheet.getRange(24, 4).getValue());

  let realAssets;
  try { realAssets = getRealAssetsSummary(); } catch(e) {}
  if (!realAssets) realAssets = { realEstate: { net: 0 }, bonds: { net: 0 }, totalNetWorthImpact: 0 };

  // --- FIX PUNTI 1 E 2: LE NUOVE DEFINIZIONI MATEMATICHE ---
  const effectiveCash = valCashEq > 0 ? valCashEq : valCash;
  
  // Liquid NW = Stocks + Crypto + Cash eq + Etfs
  const mathLiquidNW = valStocks + valCrypto + effectiveCash + valEtfs;
  
  // Total NW = LiquidNW + Others + Real Estate + Bonds + Pension + (Eventuali Debiti)
  const mathGrandTotal = mathLiquidNW + valOthers + valPension + realAssets.totalNetWorthImpact;

  // Base 100% pura per il grafico a torta (Asset Allocation senza debiti liberi)
  const mathAllocatedTotal = mathLiquidNW + valOthers + valPension + realAssets.realEstate.net + realAssets.bonds.net;

  const fmt = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
  const calcPct = (raw) => mathAllocatedTotal > 0 ? ((raw / mathAllocatedTotal) * 100).toFixed(1) + "%" : "0.0%";

  const getRecalcRow = (row, rawNum) => {
    return { amount: sheet.getRange(row, 2).getDisplayValue(), raw: rawNum, percent: calcPct(rawNum) };
  };

  const getSectionData = (startRow) => {
    return {
      unrealized: { amount: sheet.getRange(startRow, 2).getDisplayValue(), percent: sheet.getRange(startRow, 3).getDisplayValue() },
      realized: { amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), percent: sheet.getRange(startRow + 1, 3).getDisplayValue() },
      balance: { amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), percent: sheet.getRange(startRow + 2, 3).getDisplayValue() }
    };
  };

  return {
    liquidNetWorth: fmt.format(mathLiquidNW),    
    liquidNetWorthUSD: sheet.getRange(26, 2).getDisplayValue(), 
    totalNetWorth: fmt.format(mathGrandTotal), 
    totalNetWorthUSD: sheet.getRange(24, 2).getDisplayValue(), 

    summary: { 
      allocatedTotal: mathAllocatedTotal, 
      etfs: getRecalcRow(2, valEtfs),      
      stocks: getRecalcRow(3, valStocks),    
      
      // FIX PUNTO 3: Unifichiamo Cash in un solo blocco per la UI
      liquid: { amount: fmt.format(effectiveCash), raw: effectiveCash, percent: calcPct(effectiveCash) },   
      
      crypto: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) },    
      others: getRecalcRow(7, valOthers),
      pension: { amount: sheet.getRange(24, 4).getDisplayValue(), raw: valPension, percent: calcPct(valPension) },
      realEstate: { amount: fmt.format(realAssets.realEstate.net), raw: realAssets.realEstate.net, percent: calcPct(realAssets.realEstate.net) },
      bonds: { amount: fmt.format(realAssets.bonds.net), raw: realAssets.bonds.net, percent: calcPct(realAssets.bonds.net) }
    },

    cryptoSection: { main: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) }, ...getSectionData(9) },
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