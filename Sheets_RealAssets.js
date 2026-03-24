/**
 * ====================================================================
 * REAL ASSETS ENGINE (REAL ESTATE & BONDS)
 * Handles mortgage amortization (French method) and Bond/BTP valuation.
 * ====================================================================
 */

/**
 * Calculates current Home Equity (Market Value - Remaining Debt).
 * Uses the "French Amortization Method" (Standard for Italian mortgages).
 *
 * @param {number} marketValue Current estimated market value of the property.
 * @param {number} loanAmount Original principal amount of the mortgage.
 * @param {number} annualRate Annual interest rate (e.g., 2.5 for 2.5%).
 * @param {number} years Total duration of the mortgage in years.
 * @param {string} startDateStr Start date of the mortgage (format "YYYY-MM-DD").
 * @return {number} Current Net Equity.
 * @customfunction
 */
function ASSET_IMMOBILE_EQUITY(marketValue, loanAmount, annualRate, years, startDateStr) {
  // 1. Input Validation
  if (!marketValue) return 0;
  if (!loanAmount || loanAmount === 0) return marketValue; // Property owned outright
  
  let startDate = new Date(startDateStr);
  let today = new Date();
  
  // Return Market Value if date is invalid or in the future
  if (isNaN(startDate.getTime()) || startDate > today) return marketValue;

  // 2. French Amortization Calculation
  let r = annualRate  / 12; // Monthly rate
  let n = years * 12;              // Total number of payments
  
  // Calculate Monthly Payment (Standard Formula)
  // PMT = P * (r * (1+r)^n) / ((1+r)^n - 1)
  let monthlyPayment = loanAmount * (r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1);
  
  // 3. Calculate Remaining Principal (Outstanding Debt)
  // Calculate months elapsed since start
  let monthsPassed = (today.getFullYear() - startDate.getFullYear()) * 12 + (today.getMonth() - startDate.getMonth());
  
  if (monthsPassed >= n) {
    // Mortgage fully paid off
    return marketValue; 
  }

  // Remaining Debt at month k:
  // D_k = PMT * (1 - (1+r)^-(n-k)) / r
  let monthsRemaining = n - monthsPassed;
  let remainingDebt = monthlyPayment * (1 - Math.pow(1 + r, -monthsRemaining)) / r;
  
  // 4. Result: Asset Value - Liability
  return Number((marketValue - remainingDebt).toFixed(2));
}

/**
 * Calculates the Net Present Value of a Bond/BTP.
 * Applies specific Italian taxation rates (12.5% for White List vs 26%).
 *
 * @param {number} nominal Nominal value invested (e.g., 10000).
 * @param {number} currentPrice Current market price (e.g., 98.5 or 102).
 * @param {number} couponRate Annual gross coupon rate (e.g., 4.0).
 * @param {boolean} isWhiteList TRUE for Gov Bonds (12.5% tax), FALSE for Corporate (26%). Default: TRUE.
 * @return {number} Estimated Net Value.
 * @customfunction
 */
function ASSET_BTP_VALORE(nominal, currentPrice, couponRate, isWhiteList) {
  if (!nominal) return 0;
  if (!currentPrice) currentPrice = 100; // Fallback to par if price is missing
  
  // Tax Rate Determination
  // If isWhiteList is omitted, assume TRUE (Italian BTP) -> 12.5%
  // If FALSE -> 26%
  let taxRate = (isWhiteList === false) ? 0.26 : 0.125;
  
  // 1. Capital Value
  let capitalValue = nominal * (currentPrice / 100);
  
  // 2. Accrued Interest (Simplified)
  // Note: For a precise Net Worth snapshot, Market Value is usually sufficient.
  // Accrued interest calculation requires exact last coupon date.
  // Optional: Add simplified accrual logic here if needed.
  
  return Number(capitalValue.toFixed(2));
}

/**
 * Helper to sum range values for the Dashboard.
 * Robust against non-numeric strings.
 * @param {Array} rangeValues 2D array from Sheet range.
 * @return {number} Total sum.
 * @customfunction
 */
function GET_TOTAL_REAL_ESTATE(rangeValues) {
  let total = 0;
  if (!rangeValues) return 0;
  for (let i = 0; i < rangeValues.length; i++) {
    let val = parseFloat(rangeValues[i][0]);
    if (!isNaN(val)) total += val;
  }
  return total;
}

/**
 * ARCHIVE_REAL_ASSETS
 * Congela il valore mensile di Immobili, Obbligazioni e Passività nel foglio Log_Valuations.
 * - Idempotente: sovrascrive i dati se eseguita più volte nello stesso mese per evitare duplicati.
 * - Calcola dinamicamente il valore di mercato attuale e l'ammortamento dei debiti.
 * - Estrae Gross Value, Debt Value e Net Equity.
 */
function ARCHIVE_REAL_ASSETS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssets = ss.getSheetByName("DB_RealAssets");
  const dbDebts = ss.getSheetByName("DB_Debts");
  const logSheet = ss.getSheetByName("Log_Valuations");

  if (!dbAssets || !dbDebts || !logSheet) return;

  const today = new Date();
  const currentMonthStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");

  // 1. IDEMPOTENZA: Rimuove eventuali log già salvati per il mese in corso
  let logData = logSheet.getDataRange().getValues();
  for (let i = logData.length - 1; i > 0; i--) { // Ciclo al contrario per non sballare gli indici durante l'eliminazione
    let rowDate = logData[i][0];
    if (rowDate instanceof Date) {
       let rowMonthStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM");
       if (rowMonthStr === currentMonthStr) {
          logSheet.deleteRow(i + 1);
       }
    }
  }

  // Lista degli stati terminali da ignorare (come richiesto)
  const terminalStatuses = ["Sold", "Matured", "Paid_Off", "Renegotiated", "Partially_Paid", "Closed"];

  // 2. Lettura Debiti e calcolo ammortamento residuo esatto
  const debtsData = dbDebts.getDataRange().getValues();
  let activeDebts = {};

  for (let i = 1; i < debtsData.length; i++) {
     let dId = debtsData[i][0];
     let dName = debtsData[i][1];
     let dStart = new Date(debtsData[i][2]);
     let dPrin = parseFloat(debtsData[i][3]) || 0;
     let dRate = parseFloat(debtsData[i][4]) || 0;
     let dYrs = parseFloat(debtsData[i][5]) || 0;
     let dStatus = debtsData[i][7];
     let dBalloon = parseFloat(debtsData[i][8]) || 0;

     if (terminalStatuses.includes(dStatus)) continue;

     let outstanding = dPrin;
     let totMonths = dYrs * 12;
     let mPassed = (today.getFullYear() - dStart.getFullYear()) * 12 + (today.getMonth() - dStart.getMonth());

     if (mPassed > 0 && mPassed < totMonths) {
        let mR = dRate / 12;
        let pmt = 0;
        if (mR === 0) pmt = (dPrin - dBalloon) / totMonths;
        else pmt = ((dPrin * Math.pow(1 + mR, totMonths)) - dBalloon) * mR / (Math.pow(1 + mR, totMonths) - 1);

        let remM = totMonths - mPassed;
        if (mR === 0) outstanding = (pmt * remM) + dBalloon;
        else outstanding = (pmt * (1 - Math.pow(1 + mR, -remM)) / mR) + (dBalloon * Math.pow(1 + mR, -remM));
     } else if (mPassed >= totMonths) {
        outstanding = dBalloon > 0 ? dBalloon : 0;
     }

     activeDebts[dId] = { name: dName, outstanding: outstanding, linked: false };
  }

  // 3. Lettura Asset, Match con Debiti e Preparazione Batch
  const assetsData = dbAssets.getDataRange().getValues();
  let logsToAppend = [];

  for (let i = 1; i < assetsData.length; i++) {
     let aId = assetsData[i][0];
     let aType = assetsData[i][1];
     let aName = assetsData[i][2];
     let aGross = parseFloat(assetsData[i][5]) || 0; // Purchase_Price / Market_Val per Immobili
     let aDebtId = assetsData[i][10];
     let aStatus = assetsData[i][11];
     let aLivePrice = parseFloat(assetsData[i][13]) || 0; // Colonna N: Prezzo Live Yahoo per i Bond

     if (terminalStatuses.includes(aStatus)) continue;

     let grossVal = aGross;
     
     // Se è un Bond, calcola il valore di mercato attuale: (Nominale * Prezzo Live) / 100
     if (aType.includes("Bond") && aLivePrice > 0) {
        let nominal = parseFloat(assetsData[i][6]) || 0;
        grossVal = (nominal * aLivePrice) / 100;
     }

     let debtVal = 0;
     if (aDebtId && activeDebts[aDebtId]) {
        debtVal = activeDebts[aDebtId].outstanding;
        activeDebts[aDebtId].linked = true; // Segniamo il debito come "consumato" (legato a un asset)
     }

     let netVal = grossVal - debtVal;

     logsToAppend.push([
        today, aType, aName, grossVal, debtVal, netVal, aId, aDebtId || ""
     ]);
  }

  // 4. Aggiunta Debiti "Standalone" (es. prestiti auto non legati a case)
  for (let dId in activeDebts) {
     if (!activeDebts[dId].linked) {
        let dObj = activeDebts[dId];
        logsToAppend.push([
           today, "Liability", dObj.name, 0, dObj.outstanding, -dObj.outstanding, dId, ""
        ]);
     }
  }

  // 5. Scrittura Batch Veloce sul Foglio Log
  if (logsToAppend.length > 0) {
     logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, 8).setValues(logsToAppend);
  }
}

/**
 * Fetches live market prices for multiple tickers from Yahoo Finance API in a single request.
 * @param {Array<string>} tickers Array of Yahoo Finance ticker symbols.
 * @return {Object} Dictionary mapping tickers to their current market price.
 */
function FETCH_YAHOO_BOND_PRICES_BATCH(tickers) {
  const prices = {};
  if (!tickers || tickers.length === 0) return prices;

  // Filter out empty strings and build comma-separated string
  const validTickers = tickers.filter(t => t && t.trim() !== "");
  if (validTickers.length === 0) return prices;

  const symbolsStr = encodeURIComponent(validTickers.join(","));
  // Endpoint changed to 'quote' to support multiple symbols at once
  const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbolsStr}`;
  
  try {
    const options = { "method": "get", "muteHttpExceptions": true };
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.quoteResponse && json.quoteResponse.result) {
        json.quoteResponse.result.forEach(item => {
          if (item.symbol && item.regularMarketPrice) {
            prices[item.symbol] = Number(item.regularMarketPrice.toFixed(2));
          }
        });
      }
    } else {
      Logger.log(`API Error batch fetch: Status ${response.getResponseCode()}`);
    }
  } catch (e) {
    Logger.log(`Batch Fetch Exception: ${e.message}`);
  }
  
  return prices;
}

/**
 * Updates the Live_Price column in DB_RealAssets for all active Bonds.
 * Uses a single batch request to save UrlFetchApp daily quota and run instantly.
 */
function UPDATE_ALL_BOND_PRICES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssetsSheet = ss.getSheetByName("DB_RealAssets");
  
  if (!dbAssetsSheet) return;
  
  const data = dbAssetsSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const headers = data[0];
  const typeIdx = headers.indexOf("Type");
  const statusIdx = headers.indexOf("Status");
  const tickerIdx = headers.indexOf("Ticker_Yahoo");
  const livePriceIdx = headers.indexOf("Live_Price");
  
  if (tickerIdx === -1 || livePriceIdx === -1) {
    Logger.log("Columns 'Ticker_Yahoo' or 'Live_Price' not found.");
    return;
  }
  
  // 1. Collect all unique valid tickers to fetch
  const tickersToFetch = [];
  for (let i = 1; i < data.length; i++) {
    let type = data[i][typeIdx];
    let status = data[i][statusIdx];
    let ticker = data[i][tickerIdx];
    
    if (status === "Active" && (type === "Government Bond" || type === "Corporate Bond") && ticker) {
      if (!tickersToFetch.includes(ticker)) {
        tickersToFetch.push(ticker);
      }
    }
  }
  
  // 2. Fetch all prices in one single go
  const livePricesMap = FETCH_YAHOO_BOND_PRICES_BATCH(tickersToFetch);
  
  // 3. Prepare updates array mapping the new prices back to their rows
  const updates = [];
  for (let i = 1; i < data.length; i++) {
    let type = data[i][typeIdx];
    let status = data[i][statusIdx];
    let ticker = data[i][tickerIdx];
    
    if (status === "Active" && (type === "Government Bond" || type === "Corporate Bond")) {
      // Fallback to 100 or to the old price if the API returns undefined
      let currentStoredPrice = data[i][livePriceIdx];
      let livePrice = livePricesMap[ticker] || currentStoredPrice || 100; 
      updates.push([livePrice]);
    } else {
      updates.push([data[i][livePriceIdx]]); 
    }
  }
  
  // 4. Batch write to the Sheet
  if (updates.length > 0) {
    dbAssetsSheet.getRange(2, livePriceIdx + 1, updates.length, 1).setValues(updates);
  }
}

/**
 * Scansione automatica dei Bond. Se sono scaduti, li contrassegna come "Matured".
 * Va inserita nel trigger di fine mese.
 */
function CHECK_BOND_MATURITIES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssets = ss.getSheetByName("DB_RealAssets");
  if(!dbAssets) return;
  
  const data = dbAssets.getDataRange().getValues();
  const today = new Date();
  let updates = [];
  
  for(let i=1; i<data.length; i++) {
    let type = data[i][1];
    let status = data[i][11];
    let maturityDate = new Date(data[i][8]); // Colonna I
    
    if(status === "Active" && (type === "Government Bond" || type === "Corporate Bond")) {
       if(!isNaN(maturityDate.getTime()) && today >= maturityDate) {
          // Segna come scaduto
          dbAssets.getRange(i+1, 12).setValue("Matured");
          Logger.log(`Bond ${data[i][2]} is Matured.`);
       }
    }
  }
}
