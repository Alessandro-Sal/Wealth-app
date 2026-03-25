/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Implements foolproof Deep Server Cache for Crypto to prevent #N/A phantom zeroing.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  SpreadsheetApp.flush(); 

  // Funzione infallibile che riconosce se la cella è in errore o sta caricando
  const isErrorText = (text) => {
    const str = String(text).toUpperCase();
    return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
  };

  // FIX CRITICO: Estrazione intelligente. Se c'è un errore, restituisce NULL, non 0.
  const getSafeNum = (row, col) => {
    const displayVal = sheet.getRange(row, col).getDisplayValue();
    if (isErrorText(displayVal)) return null; // Segnale chiaro che la formula è rotta/in caricamento

    const val = sheet.getRange(row, col).getValue();
    if (typeof val === 'number') return val;
    const parsed = parseFloat(String(val).replace(/[^0-9,-]+/g,"").replace(',', '.'));
    return isNaN(parsed) ? 0 : parsed;
  };

  let cryptoStr = sheet.getRange(6, 2).getDisplayValue();
  let cryptoNum = getSafeNum(6, 2);
  let retries = 0;
  
  // Aspetta SOLO se la cella è palesemente in errore (null)
  while (cryptoNum === null && retries < 4) {
    Utilities.sleep(1000); 
    SpreadsheetApp.flush(); 
    cryptoStr = sheet.getRange(6, 2).getDisplayValue();
    cryptoNum = getSafeNum(6, 2);
    retries++;
  }

  const getSectionData = (startRow) => {
    return {
      unrealized: { amount: sheet.getRange(startRow, 2).getDisplayValue(), percent: sheet.getRange(startRow, 3).getDisplayValue() },
      realized: { amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), percent: sheet.getRange(startRow + 1, 3).getDisplayValue() },
      balance: { amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), percent: sheet.getRange(startRow + 2, 3).getDisplayValue() }
    };
  };

  // --- DEEP CACHE SALVAVITA CRYPTO ---
  const scriptProps = PropertiesService.getScriptProperties();
  let cryptoSectionData;

  if (cryptoNum !== null) {
      // Dati perfetti (Anche se fosse legittimamente 0€)! -> Aggiorniamo il Vault
      cryptoSectionData = getSectionData(9);
      const validCryptoState = { raw: cryptoNum, str: cryptoStr, section: cryptoSectionData };
      scriptProps.setProperty('DEEP_CACHE_CRYPTO', JSON.stringify(validCryptoState));
  } else {
      // La cella è rimasta rotta/N/A -> Peschiamo dal Vault per evitare lo zero in UI!
      const deepCache = scriptProps.getProperty('DEEP_CACHE_CRYPTO');
      if (deepCache) {
          try {
              const parsed = JSON.parse(deepCache);
              cryptoNum = parsed.raw;
              cryptoStr = parsed.str;
              cryptoSectionData = parsed.section;
          } catch(e) {}
      }
      // Se il Vault è vuoto, forziamo a 0 per non far crashare l'app
      if (cryptoNum === null) {
          cryptoNum = 0;
          cryptoStr = "€ 0,00";
          cryptoSectionData = { unrealized: { amount: "--", percent: "--" }, realized: { amount: "--", percent: "--" }, balance: { amount: "--", percent: "--" } };
      }
  }

  // --- LETTURA ASSET (I null diventano 0 automaticamente per evitare bug matematici) ---
  const valEtfs = getSafeNum(2, 2) || 0;
  const valStocks = getSafeNum(3, 2) || 0;
  const valCash = getSafeNum(4, 2) || 0;
  const valCashEq = getSafeNum(5, 2) || 0;
  const valOthers = getSafeNum(7, 2) || 0;
  const valPension = getSafeNum(24, 4) || 0;
  const valCrypto = cryptoNum; // Già validato dal Vault

  let realAssets;
  try { realAssets = getRealAssetsSummary(); } catch(e) {}
  if (!realAssets) realAssets = { realEstate: { net: 0 }, bonds: { net: 0 }, totalNetWorthImpact: 0 };

  // --- REGOLE NW E CALCOLO PERCENTUALI ---
  const effectiveCash = valCashEq > 0 ? valCashEq : valCash;
  const mathLiquidNW = valStocks + valCrypto + effectiveCash + valEtfs;
  const mathGrandTotal = mathLiquidNW + valOthers + valPension + realAssets.totalNetWorthImpact;
  const mathAllocatedTotal = mathLiquidNW + valOthers + valPension + realAssets.realEstate.net + realAssets.bonds.net;

  const fmt = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
  const calcPct = (raw) => mathAllocatedTotal > 0 ? ((raw / mathAllocatedTotal) * 100).toFixed(1) + "%" : "0.0%";

  const getRecalcRow = (row, rawNum) => {
    return { amount: sheet.getRange(row, 2).getDisplayValue(), raw: rawNum, percent: calcPct(rawNum) };
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
      liquid: { amount: fmt.format(effectiveCash), raw: effectiveCash, percent: calcPct(effectiveCash) },   
      crypto: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) },    
      others: getRecalcRow(7, valOthers),
      pension: { amount: sheet.getRange(24, 4).getDisplayValue(), raw: valPension, percent: calcPct(valPension) },
      realEstate: { amount: fmt.format(realAssets.realEstate.net), raw: realAssets.realEstate.net, percent: calcPct(realAssets.realEstate.net) },
      bonds: { amount: fmt.format(realAssets.bonds.net), raw: realAssets.bonds.net, percent: calcPct(realAssets.bonds.net) }
    },

    cryptoSection: { main: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) }, ...cryptoSectionData },
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