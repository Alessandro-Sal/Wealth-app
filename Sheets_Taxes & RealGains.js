/**
 * ====================================================================
 * MASTER SCRIPT: COMPLETE SUITE (PORTFOLIO + FISCAL + TAX CREDIT BASKET)
 * Version: FINAL FIX (Resolved Reference Errors)
 * ====================================================================
 */

// --- 1. GLOBAL CONFIGURATION ---
// Placed at the top to avoid "not defined" errors
const INCLUDE_CRYPTO_IN_TAX_BASKET = false; 


// --- 2. COMMON UTILITIES (DATA CLEANING) ---

function cleanNumber(value) {
  if (!value) return 0;
  if (typeof value === 'number') return value;
  let str = value.toString().trim();
  // If it's "x" or empty, return 0
  if (str === "" || str.toLowerCase() === "x") return 0;
  
  str = str.replace(/[€$£\s]/g, ""); 
  if (str.includes(".") && str.includes(",")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  }
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parseDate(dateVal) {
  return dateVal ? new Date(dateVal) : new Date();
}

/**
 * Normalizes raw spreadsheet data into a standardized object array.
 * Shared by all reporting engines.
 */
function normalizeData(dates, security, actions, quantity, priceCol, spentCol, typeCol, isCryptoMode) {
  let cleanData = [];
  
  const flatDates = dates ? dates.map(r => r[0]) : [];
  const flatSec = security ? security.map(r => r[0]) : [];
  const flatAct = actions ? actions.map(r => r[0]) : [];
  const flatQty = quantity ? quantity.map(r => r[0]) : [];
  
  const flatPriceUnit = priceCol ? priceCol.map(r => r[0]) : []; 
  const flatSpent = spentCol ? spentCol.map(r => r[0]) : []; 
  const flatType = typeCol ? typeCol.map(r => r[0]) : []; 

  for (let i = 0; i < flatSec.length; i++) {
    if (!flatSec[i] || flatSec[i] === "") continue;

    let ticker = flatSec[i].toString();
    // Exclude Cash rows
    if (ticker.toUpperCase().includes("CASH") || ticker.toUpperCase().includes("EUR")) continue;

    let action = flatAct[i] ? flatAct[i].toString() : "Buy";
    let qty = cleanNumber(flatQty[i]);
    let date = parseDate(flatDates[i]);
    
    // --- VALUE HANDLING ---
    let unitPrice = cleanNumber(flatPriceUnit[i]); 
    let totalSpent = 0;
    let assetType = "Stock"; 

    if (isCryptoMode) {
      assetType = "Crypto";
      let rawSpent = cleanNumber(flatSpent[i]); 
      
      // SMART LOGIC FOR CRYPTO:
      // If unit price is missing but we have Total Spent & Qty, calculate Unit Price
      if (unitPrice === 0 && qty !== 0 && rawSpent !== 0) unitPrice = rawSpent / qty;

      if (rawSpent === 0 && qty !== 0 && unitPrice !== 0) totalSpent = qty * unitPrice;
      else totalSpent = rawSpent;

    } else {
      // STOCK LOGIC
      if (flatType.length > i && flatType[i]) assetType = flatType[i].toString();
      
      totalSpent = unitPrice * qty;
      
      // Handle Dividends (spent column = dividend amount)
      if (action === "Dividend" || action === "Dividends") {
         totalSpent = unitPrice; 
         qty = 0; 
      }
    }

    cleanData.push({
      date: date,
      ticker: ticker,
      action: action,
      qty: Number(qty),
      unitPrice: Number(unitPrice),
      totalSpent: Number(totalSpent),
      type: assetType
    });
  }
  
  cleanData.sort((a, b) => a.date - b.date);
  return cleanData;
}

// --- 3. ENGINE A: PORTFOLIO DASHBOARD (WITH SELL METRICS) ---

function calculatePortfolioStats(allData) {
  let portfolio = new Map(); 
  let stats = new Map();     
  
  // Sort chronologically
  allData.sort((a, b) => a.date - b.date);

  // --- PHASE 1: PROCESS TRANSACTIONS ---
  // 
  for (let trade of allData) {
    let ticker = trade.ticker;
    let precision = (trade.type === "Crypto") ? 8 : 5;

    // Initialize Stats Object
    if (!stats.has(ticker)) {
      stats.set(ticker, { 
        type: trade.type, 
        tradingPnL: 0, 
        dividends: 0, 
        totalInvestedHistorical: 0,
        // Buy Metrics
        firstBuyDate: null,
        firstBuyPrice: 0,
        minBuyPrice: Infinity, 
        maxBuyPrice: 0,
        // Sell Metrics (NEW)
        maxSellPrice: 0,
        totalSoldShares: 0,
        totalSoldRevenue: 0,
        // Behavioral Metrics
        tradeCount: 0,
        lastActionDate: null
      });
    }
    let s = stats.get(ticker);

    s.lastActionDate = trade.date;
    s.tradeCount += 1;

    // Dividend Handling
    if (trade.action === "Dividend" || trade.action === "Dividends") {
      s.dividends += trade.totalSpent;
      continue;
    }

    // Buy Handling
    if (["Buy", "DRIP", "REWARD", "STAKING", "MINING", "MINT"].includes(trade.action) || trade.action.toUpperCase().includes("BUY")) {
      if (s.totalInvestedHistorical === 0) {
        s.firstBuyDate = trade.date;
        s.firstBuyPrice = trade.unitPrice;
      }
      if (trade.unitPrice > 0) {
        if (trade.unitPrice < s.minBuyPrice) s.minBuyPrice = trade.unitPrice;
        if (trade.unitPrice > s.maxBuyPrice) s.maxBuyPrice = trade.unitPrice;
      }
      s.totalInvestedHistorical += trade.totalSpent;
      
      let lot = { shares: trade.qty, price: trade.unitPrice, date: trade.date };
      if (!portfolio.has(ticker)) portfolio.set(ticker, [lot]);
      else portfolio.get(ticker).push(lot);
    }
    // Sell Handling
    else if (trade.action === "Sell") {
      
      // --- SELL METRICS LOGIC (NEW) ---
      if (trade.unitPrice > s.maxSellPrice) s.maxSellPrice = trade.unitPrice;
      s.totalSoldShares += trade.qty;
      s.totalSoldRevenue += (trade.qty * trade.unitPrice);
      // ----------------------------------------

      if (!portfolio.has(ticker)) continue;
      let activeTrades = portfolio.get(ticker);
      let sharesToSell = Number(trade.qty.toFixed(precision));
      let salePrice = trade.unitPrice;
      
      // FIFO Loop to close lots
      while (sharesToSell > 0 && activeTrades.length > 0) {
        let oldestTrade = activeTrades[0];
        let availableShares = Number(oldestTrade.shares.toFixed(precision));
        let sharesTaken = (availableShares <= sharesToSell) ? availableShares : sharesToSell;
        
        oldestTrade.shares = Number((oldestTrade.shares - sharesTaken).toFixed(precision));
        sharesToSell = Number((sharesToSell - sharesTaken).toFixed(precision));
        
        if (oldestTrade.shares === 0) activeTrades.shift();

        let costBasis = sharesTaken * oldestTrade.price;
        let revenue = sharesTaken * salePrice;
        s.tradingPnL += (revenue - costBasis);
      }
      if (activeTrades.length === 0) portfolio.delete(ticker);
    }
  }

  // --- PHASE 2: CALCULATE TOTALS ---
  let allTickers = Array.from(stats.keys()).sort();
  let grandTotalBookValue = 0;
  let tickerBookValues = new Map();

  for (let ticker of allTickers) {
    let currentShares = 0;
    let currentBookValue = 0;
    let oldestActiveDate = null;

    if (portfolio.has(ticker)) {
      let lots = portfolio.get(ticker);
      if (lots.length > 0) oldestActiveDate = lots[0].date;
      for (let lot of lots) {
        currentShares += lot.shares;
        currentBookValue += (lot.shares * lot.price);
      }
    }
    tickerBookValues.set(ticker, { shares: currentShares, bookVal: currentBookValue, activeSince: oldestActiveDate });
    grandTotalBookValue += currentBookValue;
  }

  // --- PHASE 3: DATA OUTPUT & ANALYSIS ---
  let output = [];
  let today = new Date();

  for (let ticker of allTickers) {
    let s = stats.get(ticker);
    let tv = tickerBookValues.get(ticker);

    let avgPrice = tv.shares > 0 ? tv.bookVal / tv.shares : 0;
    let status = (tv.shares > 0.00000001) ? "OPEN" : "CLOSED";
    let totalRealizedPnL = s.tradingPnL + s.dividends;
    
    let breakEven = 0;
    if (tv.shares > 0) {
      breakEven = (tv.bookVal - totalRealizedPnL) / tv.shares;
    }

    let allocationPct = (grandTotalBookValue > 0) ? (tv.bookVal / grandTotalBookValue) : 0;
    let tradingRoiPct = (s.totalInvestedHistorical > 0) ? (s.tradingPnL / s.totalInvestedHistorical) : 0;
    let totalRoiPct = (s.totalInvestedHistorical > 0) ? (totalRealizedPnL / s.totalInvestedHistorical) : 0;

    let daysHeld = 0;
    if (status === "OPEN" && tv.activeSince) {
      let timeDiff = today.getTime() - tv.activeSince.getTime();
      daysHeld = Math.floor(timeDiff / (1000 * 3600 * 24));
    }

    let finalMinPrice = (s.minBuyPrice === Infinity) ? 0 : s.minBuyPrice;
    
    // Calculate Average Sell Price
    let avgSellPrice = (s.totalSoldShares > 0) ? (s.totalSoldRevenue / s.totalSoldShares) : 0;

    // --- AUTO-ANALYSIS LOGIC ---
    let analysisNotes = [];

    if (tradingRoiPct < -0.01 && totalRoiPct > 0) analysisNotes.push("🐮 Cash Cow: Gain via Divs.");
    if (tradingRoiPct > 0.05 && (totalRoiPct - tradingRoiPct) < 0.01) analysisNotes.push("📈 Pure Trading.");
    if (finalMinPrice > 0 && finalMinPrice < (s.firstBuyPrice * 0.85)) analysisNotes.push("📉 Good DCA.");
    if (status === "OPEN" && avgPrice > (s.maxBuyPrice * 0.95) && s.tradeCount > 1) analysisNotes.push("⚠️ High Load Price.");
    if (status === "OPEN" && daysHeld > 730 && totalRoiPct < 0) analysisNotes.push("💤 Dead Money.");
    
    // New Note: Selling higher than buy max
    if (avgSellPrice > s.maxBuyPrice && avgSellPrice > 0) analysisNotes.push("💰 Sniper: Selling higher than buying.");

    let finalNote = analysisNotes.join(" | ");

    output.push([
      ticker, 
      s.type, 
      status, 
      tv.shares, 
      avgPrice, 
      s.totalInvestedHistorical, 
      s.tradingPnL, 
      s.dividends, 
      totalRealizedPnL, 
      breakEven,
      tv.bookVal,     
      allocationPct,  
      tradingRoiPct,  
      totalRoiPct,    
      s.firstBuyDate, 
      s.firstBuyPrice,
      finalMinPrice,  
      s.maxBuyPrice,
      // NEW COLUMNS
      s.maxSellPrice, // Col S
      avgSellPrice,   // Col T
      // -----------
      daysHeld,       // U
      s.lastActionDate,// V
      s.tradeCount,   // W
      finalNote       // X
    ]);
  }
  return output;
}


// --- 4. ENGINE B: DETAILED FISCAL REPORT (Reference Error Fix) ---
// Renamed from calculateFiscalReport to calculateFiscalReportDetails for consistency

function calculateFiscalReportDetails(allData) {
  let portfolio = new Map();
  let report = [];
  report.push(["Date", "Year", "Ticker", "Type", "Sold Qty", "Sell Price", "Load Price", "P&L (Gain)", "Tax (26%)", "Capital Loss"]);
  
  // Sort by date
  allData.sort((a, b) => a.date - b.date);

  for (let trade of allData) {
    let ticker = trade.ticker;
    let precision = (trade.type === "Crypto") ? 8 : 5;

    if (["Buy", "DRIP", "REWARD", "STAKING", "MINING", "MINT"].includes(trade.action) || trade.action.toUpperCase().includes("BUY")) {
      let lot = { shares: trade.qty, price: trade.unitPrice };
      if (!portfolio.has(ticker)) portfolio.set(ticker, [lot]);
      else portfolio.get(ticker).push(lot);
    }
    else if (trade.action === "Sell") {
      if (!portfolio.has(ticker)) continue;
      let activeTrades = portfolio.get(ticker);
      let sharesToSell = Number(trade.qty.toFixed(precision));
      let salePrice = trade.unitPrice;
      let totalCost = 0;
      let sharesSold = 0;

      while (sharesToSell > 0 && activeTrades.length > 0) {
        let oldestTrade = activeTrades[0];
        let availableShares = Number(oldestTrade.shares.toFixed(precision));
        let sharesTaken = (availableShares <= sharesToSell) ? availableShares : sharesToSell;
        
        oldestTrade.shares = Number((oldestTrade.shares - sharesTaken).toFixed(precision));
        sharesToSell = Number((sharesToSell - sharesTaken).toFixed(precision));
        
        if (oldestTrade.shares === 0) activeTrades.shift();
        totalCost += (sharesTaken * oldestTrade.price);
        sharesSold += sharesTaken;
      }
      
      if (sharesSold > 0) {
        let revenue = sharesSold * salePrice;
        let pnl = revenue - totalCost;
        let tax = (pnl > 0) ? pnl * 0.26 : 0;
        let minus = (pnl < 0) ? Math.abs(pnl) : 0;
        let year = Utilities.formatDate(trade.date, Session.getScriptTimeZone(), "yyyy");
        let avgBuy = totalCost / sharesSold;

        report.push([
          trade.date, year, ticker, trade.type, sharesSold, salePrice, avgBuy, pnl, tax, minus
        ]);
      }
      if (activeTrades.length === 0) portfolio.delete(ticker);
    }
  }
  return report;
}


// --- 5. ENGINE C: TAX CREDIT BASKET ("ZAINETTO") WITH STRATEGY ---

function calculateZainettoReport(allData) {
  // 1. Calculate Detailed Sales (Function now exists!)
  let trades = calculateFiscalReportDetails(allData); 
  
  // 2. Retrieve Dividends
  let yearlyDividends = new Map();
  for (let d of allData) {
     if (d.action === "Dividend" || d.action === "Dividends") {
       let y = Utilities.formatDate(d.date, Session.getScriptTimeZone(), "yyyy");
       yearlyDividends.set(y, (yearlyDividends.get(y) || 0) + d.totalSpent);
     }
  }

  trades.shift(); // Remove header row
  let yearlyStats = new Map();

  // 3. Aggregate Annual P&L
  for (let t of trades) {
    let year = parseInt(t[1]); 
    let type = t[3];     
    let pnl = t[7];      

    if (!yearlyStats.has(year)) {
      yearlyStats.set(year, { gainCompensabile: 0, gainNonCompensabile: 0, lossAnno: 0 });
    }
    let s = yearlyStats.get(year);

    if (pnl >= 0) {
      if (type === "ETF") s.gainNonCompensabile += pnl; // ETFs are usually not compensable in Italy
      else s.gainCompensabile += pnl; 
    } else {
      s.lossAnno += Math.abs(pnl);
    }
  }

  // Add years that only have dividends
  for (let yearStr of yearlyDividends.keys()) {
    let year = parseInt(yearStr);
    if (!yearlyStats.has(year)) {
       yearlyStats.set(year, { gainCompensabile: 0, gainNonCompensabile: 0, lossAnno: 0 });
    }
  }

  let report = [];
  report.push([
    "Year", "Useful Gain (Stock)", "Useless Gain (ETF/Div)", "New Losses", 
    "Basket USED", "Basket EXPIRED", "Net Taxable", 
    "Tax (26%)", "Residual Basket", "Analysis & Strategy"
  ]);

  let years = Array.from(yearlyStats.keys()).sort((a, b) => a - b);
  let minusBucket = new Map(); 

  for (let year of years) {
    let s = yearlyStats.get(year);
    let divs = yearlyDividends.get(year.toString()) || 0;
    let gainNoComp = s.gainNonCompensabile + divs; // Always Taxed
    let tradingGain = s.gainCompensabile;          // Compensable
    
    let expiredAmount = 0;
    let expiryYearLimit = year - 4; 

    // A. USE BASKET
    let usedZainetto = 0;
    let taxableBase = 0;

    if (tradingGain > 0) {
      let availableYears = Array.from(minusBucket.keys()).sort((a, b) => a - b);
      let remainingGain = tradingGain;

      for (let pastYear of availableYears) {
        if (year - pastYear > 4) continue; 
        let available = minusBucket.get(pastYear);
        let take = Math.min(available, remainingGain);
        
        minusBucket.set(pastYear, available - take);
        usedZainetto += take;
        remainingGain -= take;
        
        if (minusBucket.get(pastYear) === 0) minusBucket.delete(pastYear);
        if (remainingGain === 0) break;
      }
      taxableBase = gainNoComp + remainingGain;
    } else {
      taxableBase = gainNoComp;
    }

    // B. NEW LOSSES
    if (s.lossAnno > 0) {
      let netAnnual = s.gainCompensabile - s.lossAnno;
      if (netAnnual < 0) {
        minusBucket.set(year, Math.abs(netAnnual));
        usedZainetto = 0; 
        taxableBase = gainNoComp; 
      }
    }

    // C. CLEAN EXPIRED LOSSES
    if (minusBucket.has(expiryYearLimit)) {
      expiredAmount = minusBucket.get(expiryYearLimit);
      minusBucket.delete(expiryYearLimit); 
    }

    // D. STRATEGIC ANALYSIS
    let totalResiduo = 0;
    let notes = [];
    let nextExpiryAmount = 0;
    
    for (let [yOrigin, amount] of minusBucket.entries()) {
      totalResiduo += amount;
      if (4 - (year - yOrigin) === 0) nextExpiryAmount += amount;
    }

    if (nextExpiryAmount > 0) notes.push(`🔴 URGENT: ${nextExpiryAmount.toFixed(0)}€ expire end of year! Sell Stocks in gain.`);
    if (gainNoComp > 0 && totalResiduo > 0) notes.push(`⚠️ INEFFICIENT: Paying tax on ETF/Div while holding losses.`);
    else if (gainNoComp > 0) notes.push(`ℹ️ Tax on ETF/Div.`);
    if (expiredAmount > 0) notes.push(`☠️ EXPIRED: ${expiredAmount.toFixed(0)}€ lost forever.`);
    if (usedZainetto > 0) notes.push(`✅ Recovered: saved ${(usedZainetto * 0.26).toFixed(0)}€ in taxes.`);
    
    if (notes.length === 0) {
      if (totalResiduo > 0) notes.push(`Basket active (${totalResiduo.toFixed(0)}€).`);
      else notes.push(`No residual losses.`);
    }
    
    if (!INCLUDE_CRYPTO_IN_TAX_BASKET) notes.push("(Crypto excluded).");

    let estimatedTax = taxableBase * 0.26;
    report.push([
      year, s.gainCompensabile, gainNoComp, 
      (s.lossAnno > s.gainCompensabile) ? (s.lossAnno - s.gainCompensabile) : 0, 
      usedZainetto, expiredAmount, taxableBase, estimatedTax, totalResiduo, notes.join(" ")
    ]);
  }
  return report;
}


// --- 6. PUBLIC FUNCTIONS (For Google Sheet) ---

/**
 * PORTFOLIO DASHBOARD
 * Always shows everything (Stocks + Crypto)
 * @customfunction
 */
function PORTFOLIO_DASHBOARD(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent) {
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true);
  return calculatePortfolioStats(stockData.concat(cryptoData));
}

/**
 * FISCAL REPORT (Sales List)
 * Always shows everything for full archive.
 * @customfunction
 */
function REPORT_FISCALE(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent) {
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true);
  // Calling the renamed function
  return calculateFiscalReportDetails(stockData.concat(cryptoData));
}

/**
 * ZAINETTO REPORT (Tax Calculation & Expiry)
 * Excludes Crypto if INCLUDE_CRYPTO_IN_TAX_BASKET = false
 * @customfunction
 */
function REPORT_ZAINETTO(sD, sS, sA, sQ, sP, sT, cD, cS, cA, cQ, cP, cSp) {
  let sData = normalizeData(sD, sS, sA, sQ, sP, null, sT, false);
  let cData = [];
  
  if (INCLUDE_CRYPTO_IN_TAX_BASKET) {
    cData = normalizeData(cD, cS, cA, cQ, cP, cSp, null, true);
  }
  
  return calculateZainettoReport(sData.concat(cData));
}

// --- 7. NEW FUNCTION: PORTFOLIO EVOLUTION YEAR BY YEAR ---

/**
 * Calculates portfolio composition at Dec 31st of each year.
 * Returns: Year | Data | Ticker | Quantity Owned
 * @customfunction
 */
function PORTFOLIO_EVOLUZIONE(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent) {
  // 1. Merge and Normalize Data (Stock + Crypto)
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true);
  let allData = stockData.concat(cryptoData);

  // 2. Find Year Range (First buy to Today)
  if (allData.length === 0) return [["No data found"]];
  
  allData.sort((a, b) => a.date - b.date);
  
  let startYear = parseInt(Utilities.formatDate(allData[0].date, Session.getScriptTimeZone(), "yyyy"));
  let endYear = parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy"));
  
  let output = [];
  output.push(["Year", "Snapshot Date", "Ticker", "Type", "Owned Qty", "Hist. Invested (at 31/12)"]);

  // 3. "Time Machine" Loop: calculate state at 31/12 for each year
  for (let year = startYear; year <= endYear; year++) {
    
    // Cutoff Date: Dec 31st of loop year
    let cutoffDate = new Date(year, 11, 31, 23, 59, 59);
    
    // If current year, use today's date for current data
    let today = new Date();
    if (year === endYear) cutoffDate = today;

    // Filter transactions happening BEFORE or ON that date
    let historicMoves = allData.filter(d => d.date <= cutoffDate);

    // Calculate holdings at that moment
    let holdings = new Map(); // Ticker -> {qty, invested}

    for (let trade of historicMoves) {
       let t = trade.ticker;
       if (!holdings.has(t)) holdings.set(t, { qty: 0, invested: 0, type: trade.type });
       let h = holdings.get(t);

       // Simple Accumulation Logic
       if (["Buy", "DRIP", "REWARD", "STAKING", "MINING", "MINT"].includes(trade.action) || trade.action.toUpperCase().includes("BUY")) {
          h.qty += trade.qty;
          h.invested += trade.totalSpent;
       } 
       else if (trade.action === "Sell") {
          // Reduce invested capital proportionally on sale
          if (h.qty > 0) {
             let avgCost = h.invested / h.qty;
             h.invested -= (trade.qty * avgCost); 
             h.qty -= trade.qty;
          }
       }
       else if (trade.action === "Split") {
          // Basic split logic: qty increases, cost stays same
          h.qty += trade.qty; 
       }
    }

    // 4. Write Rows for this Year
    for (let [ticker, data] of holdings) {
       // Filter out near-zero residuals
       if (data.qty > 0.000001) {
          output.push([
             year, 
             cutoffDate, 
             ticker, 
             data.type,
             data.qty, 
             data.invested // Net spent for those shares at that date
          ]);
       }
    }
  }

  return output;
}


// --- 8. NEW FUNCTION: CASH FLOWS (Split by STOCK and ETF) ---

/**
 * Calculates yearly cash movements, separating STOCK and ETF.
 * @customfunction
 */
function PORTFOLIO_FLUSSI(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent) {
  // 1. Merge and Normalize
  // Note: sType is passed here and becomes trade.type inside allData
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true);
  let allData = stockData.concat(cryptoData);

  // Complex Map: Key = "Year_Type"
  let flows = new Map(); 

  // 2. Analyze each transaction
  for (let trade of allData) {
    if (!trade.date) continue;
    
    let year = Utilities.formatDate(trade.date, Session.getScriptTimeZone(), "yyyy");
    
    // Retrieve Type (Stock, ETF, Crypto). Default to Stock if empty.
    let type = trade.type || "Stock"; 

    // Create unique key for grouping
    let key = year + "_" + type;

    if (!flows.has(key)) {
      flows.set(key, { year: year, type: type, bought: 0, sold: 0, dividends: 0 });
    }
    let f = flows.get(key);

    // CASH FLOW LOGIC
    // A. Outflows (Buys)
    if (["Buy", "DRIP", "REWARD", "STAKING", "MINING", "MINT"].includes(trade.action) || trade.action.toUpperCase().includes("BUY")) {
       f.bought += trade.totalSpent;
    }
    // B. Inflows (Sales)
    else if (trade.action === "Sell") {
       let revenue = (trade.type === "Crypto" && trade.totalSpent > 0) ? trade.totalSpent : (trade.qty * trade.unitPrice);
       f.sold += revenue;
    }
    // C. Dividends
    else if (trade.action === "Dividend" || trade.action === "Dividends") {
       f.dividends += trade.totalSpent;
    }
  }

  // 3. Tabular Output
  let output = [];
  output.push([
    "Year", 
    "Asset Type", 
    "Deposits (Buys)", 
    "Withdrawals (Sells)", 
    "Net Dividends", 
    "Total Inflow", 
    "Net Flow"
  ]);

  // Sort by Year then Type
  let sortedKeys = Array.from(flows.keys()).sort();

  for (let key of sortedKeys) {
    let f = flows.get(key);
    let totalCashIn = f.sold + f.dividends; 
    let netFlow = totalCashIn - f.bought;

    // Print only if there was movement
    if (f.bought !== 0 || totalCashIn !== 0) {
      output.push([
        f.year,
        f.type,       // NEW COLUMN: Stock or ETF
        f.bought,     
        f.sold,       
        f.dividends,  
        totalCashIn,  
        netFlow       
      ]);
    }
  }

  return output;
}