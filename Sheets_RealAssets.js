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
 * Automates the periodic snapshot for Real Assets and Debts.
 * Reads active items from DB_RealAssets and DB_Debts, calculates current equity,
 * and appends new rows to the Log_Valuations sheet.
 */
function ARCHIVE_REAL_ASSETS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssetsSheet = ss.getSheetByName("DB_RealAssets");
  const dbDebtsSheet = ss.getSheetByName("DB_Debts");
  const logSheet = ss.getSheetByName("Log_Valuations");

  if (!dbAssetsSheet || !dbDebtsSheet || !logSheet) {
    Logger.log("Error: One or more Real Assets database sheets are missing.");
    return;
  }

  // Retrieve data ranges
  const assetsData = dbAssetsSheet.getDataRange().getValues();
  const debtsData = dbDebtsSheet.getDataRange().getValues();
  
  const assetsHeaders = assetsData[0];
  const debtsHeaders = debtsData[0];
  
  // Create a dictionary of active debts for fast lookup by Debt_ID
  const activeDebts = {};
  for (let i = 1; i < debtsData.length; i++) {
    let row = debtsData[i];
    let status = row[debtsHeaders.indexOf("Status")];
    if (status === "Active") {
      activeDebts[row[debtsHeaders.indexOf("Debt_ID")]] = {
        startDate: row[debtsHeaders.indexOf("Start_Date")],
        amount: row[debtsHeaders.indexOf("Original_Amount")],
        rate: row[debtsHeaders.indexOf("Annual_Rate_%")],
        years: row[debtsHeaders.indexOf("Duration_Years")]
      };
    }
  }

  const today = new Date();
  const logDateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const yearMonthStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");
  
  const newLogRows = [];

  // Iterate through assets to generate valuation logs
  for (let i = 1; i < assetsData.length; i++) {
    let row = assetsData[i];
    let status = row[assetsHeaders.indexOf("Status")];
    
    if (status !== "Active") continue;

    let assetId = row[assetsHeaders.indexOf("Asset_ID")];
    let type = row[assetsHeaders.indexOf("Type")];
    let purchasePrice = row[assetsHeaders.indexOf("Purchase_Price")];
    
    // Fallback: using Purchase Price as current Market Value if no external API is connected yet
    let marketValue = purchasePrice; 
    let outstandingDebt = "";
    let netEquity = 0;

    if (type === "Real Estate") {
      let linkedDebtId = row[assetsHeaders.indexOf("Linked_Debt_ID")];
      
      if (linkedDebtId && activeDebts[linkedDebtId]) {
        let debt = activeDebts[linkedDebtId];
        
        // Calculate equity using the existing French Amortization function
        netEquity = ASSET_IMMOBILE_EQUITY(marketValue, debt.amount, debt.rate, debt.years, debt.startDate);
        
        // Derive outstanding debt 
        outstandingDebt = Number((marketValue - netEquity).toFixed(2)); 
      } else {
        // Property owned outright
        netEquity = marketValue; 
      }
      
      newLogRows.push([logDateStr, yearMonthStr, assetId, marketValue, outstandingDebt, netEquity, "Automated Monthly Snapshot"]);
      
    } else if (type === "Government Bond" || type === "Corporate Bond") {
      let nominal = row[assetsHeaders.indexOf("Nominal_Value")];
      let isWhiteList = (row[assetsHeaders.indexOf("Tax_Status")] === "WhiteList");
      let couponRate = row[assetsHeaders.indexOf("Coupon_Rate_%")];
      
      // Retrieve the freshly updated live price from the dedicated column
      let livePriceIdx = assetsHeaders.indexOf("Live_Price");
      let currentPrice = row[livePriceIdx];
      
      // Fallback in case the column is empty
      if (!currentPrice || currentPrice === "") {
         currentPrice = 100; 
      }
      
      netEquity = ASSET_BTP_VALORE(nominal, currentPrice, couponRate, isWhiteList);
      
      newLogRows.push([logDateStr, yearMonthStr, assetId, currentPrice, "", netEquity, "Automated Monthly Snapshot"]);
    }
  }

  // Batch append all generated rows to Log_Valuations
  if (newLogRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogRows.length, newLogRows[0].length).setValues(newLogRows);
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
