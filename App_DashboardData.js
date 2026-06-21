/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Implements foolproof Deep Server Cache for Crypto to prevent #N/A phantom zeroing.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  // FIX (1.3): UN'UNICA lettura batch del blocco A1:D26 invece di ~39 getRange(...) singoli.
  // Ogni getValue()/getDisplayValue() su singola cella e' un round-trip ai server Sheets;
  // prima la stessa cella veniva letta anche 2 volte (display + value). Ora 2 letture totali.
  // Rimosso anche il SpreadsheetApp.flush() iniziale (inutile in sola lettura).
  const NUM_ROWS = 26;
  const NUM_COLS = 4;
  const rawVals  = sheet.getRange(1, 1, NUM_ROWS, NUM_COLS).getValues();
  const dispVals = sheet.getRange(1, 1, NUM_ROWS, NUM_COLS).getDisplayValues();
  const cellRaw  = (row, col) => rawVals[row - 1][col - 1];
  const cellDisp = (row, col) => dispVals[row - 1][col - 1];

  // Funzione infallibile che riconosce se la cella è in errore o sta caricando
  const isErrorText = (text) => {
    const str = String(text).toUpperCase();
    return str === "" || str.includes("#") || str.includes("LOAD") || str.includes("ERROR") || str.includes("N/A");
  };

  // FIX CRITICO: Estrazione intelligente. Se c'è un errore, restituisce NULL, non 0.
  const getSafeNum = (row, col) => {
    const displayVal = cellDisp(row, col);
    if (isErrorText(displayVal)) return null; // Segnale chiaro che la formula è rotta/in caricamento

    const val = cellRaw(row, col);
    if (typeof val === 'number') return val;
    const parsed = parseFloat(String(val).replace(/[^0-9,-]+/g,"").replace(',', '.'));
    return isNaN(parsed) ? 0 : parsed;
  };

  // FIX (1.3): rimosso il loop di retry con Utilities.sleep(1000) x4 (fino a 4s di BLOCCO
  // propagati al primo round-trip). Se la cella crypto e' in errore/caricamento, si usa
  // direttamente il "Vault" DEEP_CACHE_CRYPTO qui sotto: stesso identico fallback gia'
  // previsto dal codice originale quando cryptoNum risultava null.
  let cryptoStr = cellDisp(6, 2);
  let cryptoNum = getSafeNum(6, 2);

  const getSectionData = (startRow) => {
    return {
      unrealized: { amount: cellDisp(startRow, 2), percent: cellDisp(startRow, 3) },
      realized: { amount: cellDisp(startRow + 1, 2), percent: cellDisp(startRow + 1, 3) },
      balance: { amount: cellDisp(startRow + 2, 2), percent: cellDisp(startRow + 2, 3) }
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
    return { amount: cellDisp(row, 2), raw: rawNum, percent: calcPct(rawNum) };
  };

  return {
    liquidNetWorth: fmt.format(mathLiquidNW),
    liquidNetWorthUSD: cellDisp(26, 2),
    totalNetWorth: fmt.format(mathGrandTotal),
    totalNetWorthUSD: cellDisp(24, 2),

    summary: { 
      allocatedTotal: mathAllocatedTotal, 
      etfs: getRecalcRow(2, valEtfs),      
      stocks: getRecalcRow(3, valStocks),    
      liquid: { amount: fmt.format(effectiveCash), raw: effectiveCash, percent: calcPct(effectiveCash) },   
      crypto: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) },    
      others: getRecalcRow(7, valOthers),
      pension: { amount: cellDisp(24, 4), raw: valPension, percent: calcPct(valPension) },
      realEstate: { amount: fmt.format(realAssets.realEstate.net), raw: realAssets.realEstate.net, percent: calcPct(realAssets.realEstate.net) },
      bonds: { amount: fmt.format(realAssets.bonds.net), raw: realAssets.bonds.net, percent: calcPct(realAssets.bonds.net) }
    },

    cryptoSection: { main: { amount: cryptoStr, raw: valCrypto, percent: calcPct(valCrypto) }, ...cryptoSectionData },
    stocksSection: { main: getRecalcRow(3, valStocks), ...getSectionData(14) },
    etfSection: { main: getRecalcRow(2, valEtfs), ...getSectionData(19) }
  };
}

/**
 * Retrieves the summary of Real Assets and Liabilities from the latest recorded month in the Log.
 * Extracts individual asset details and their exact Type for grouping.
 */
function getRealAssetsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log_Valuations");
  if (!logSheet) return null;

  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  // 1. Find the latest recorded month
  let latestDate = new Date(0);
  for (let i = 1; i < data.length; i++) {
     let d = new Date(data[i][0]);
     if (!isNaN(d.getTime()) && d > latestDate) latestDate = d;
  }
  let latestMonthStr = Utilities.formatDate(latestDate, Session.getScriptTimeZone(), "yyyy-MM");

  // 2. Prepare the summary object with item arrays
  let summary = {
     realEstate: { gross: 0, debt: 0, net: 0, items: [] },
     bonds: { gross: 0, debt: 0, net: 0, items: [] },
     liabilities: { gross: 0, debt: 0, net: 0, items: [] }, 
     totalNetWorthImpact: 0
  };

  // 3. Sum data for the latest month and collect individual items with their Type
  for (let i = 1; i < data.length; i++) {
     let rowDate = new Date(data[i][0]);
     if (isNaN(rowDate.getTime())) continue;
     
     if (Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM") === latestMonthStr) {
        let type = String(data[i][1]).trim(); // Column B: Exact type (e.g., "Corporate Bond")
        let name = data[i][2] || "Unknown";   // Column C: Asset Name
        let gross = parseFloat(data[i][3]) || 0;
        let debt = parseFloat(data[i][4]) || 0;
        let net = parseFloat(data[i][5]) || 0;

        // Use includes() to catch composite types like "Real Estate - Residential"
        if (type.includes("Real Estate")) {
           summary.realEstate.gross += gross;
           summary.realEstate.debt += debt;
           summary.realEstate.net += net;
           summary.realEstate.items.push({ name: name, type: type, net: net, debt: debt });
           
        } else if (type.includes("Bond")) {
           summary.bonds.gross += gross;
           summary.bonds.debt += debt;
           summary.bonds.net += net;
           summary.bonds.items.push({ name: name, type: type, net: net });
           
        // Catch Liability, Loan, or Mortgage keywords
        } else if (type.includes("Liability") || type.includes("Loan") || type.includes("Mortgage")) {
           summary.liabilities.gross += gross;
           summary.liabilities.debt += debt;
           summary.liabilities.net += net;
           summary.liabilities.items.push({ name: name, type: type, net: net });
        }
        
        // Calculate total impact on net worth
        summary.totalNetWorthImpact += net;
     }
  }
  
  return summary;
}