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
 * @param {Boolean} keepCash - If true, prevents filtering out "Cash" or "EUR" rows.
 */
function normalizeData(dates, security, actions, quantity, priceCol, spentCol, typeCol, isCryptoMode, keepCash = false) {
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
    
    // SKIP CASH ROWS ONLY IF keepCash IS FALSE
    if (!keepCash && (ticker.toUpperCase().includes("CASH") || ticker.toUpperCase().includes("EUR"))) continue;

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
      
      if (unitPrice === 0 && qty !== 0 && rawSpent !== 0) unitPrice = rawSpent / qty;

      if (rawSpent === 0 && qty !== 0 && unitPrice !== 0) totalSpent = qty * unitPrice;
      else totalSpent = rawSpent;

    } else {
      if (flatType.length > i && flatType[i]) assetType = flatType[i].toString();
      
      totalSpent = unitPrice * qty;
      
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

// --- 3. ENGINE A: PORTFOLIO DASHBOARD (WITH METRICS, XIRR, TWR & EFFICIENCY) ---

function calculatePortfolioStats(allData, currentPrices = new Map()) {
  let portfolio = new Map(); 
  let stats = new Map();     
  
  // Sort chronologically
  allData.sort((a, b) => a.date - b.date);

  // --- PHASE 1: PROCESS TRANSACTIONS ---
  for (let trade of allData) {
    let ticker = trade.ticker;
    let precision = (trade.type === "Crypto") ? 8 : 5;

    // Initialize Stats Object with extended efficiency metrics
    if (!stats.has(ticker)) {
      stats.set(ticker, { 
        type: trade.type, 
        tradingPnL: 0, 
        dividends: 0, 
        totalInvestedHistorical: 0,
        firstBuyDate: null,
        firstBuyPrice: 0,
        minBuyPrice: Infinity, 
        maxBuyPrice: 0,
        maxSellPrice: 0,
        totalSoldShares: 0,
        totalSoldRevenue: 0,
        tradeCount: 0,
        lastActionDate: null,
        cashFlows: [], 
        winningTrades: 0, // Efficiency Metric
        losingTrades: 0,  // Efficiency Metric
        grossProfit: 0,   // Efficiency Metric
        grossLoss: 0      // Efficiency Metric
      });
    }
    let s = stats.get(ticker);

    s.lastActionDate = trade.date;
    s.tradeCount += 1;

    // Dividend Handling
    if (trade.action === "Dividend" || trade.action === "Dividends") {
      s.dividends += trade.totalSpent;
      s.cashFlows.push({ date: trade.date, amount: trade.totalSpent }); 
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
      
      s.cashFlows.push({ date: trade.date, amount: -trade.totalSpent }); 
      
      let lot = { shares: trade.qty, price: trade.unitPrice, date: trade.date };
      if (!portfolio.has(ticker)) portfolio.set(ticker, [lot]);
      else portfolio.get(ticker).push(lot);
    }
    // Sell Handling
    else if (trade.action === "Sell") {
      if (trade.unitPrice > s.maxSellPrice) s.maxSellPrice = trade.unitPrice;
      s.totalSoldShares += trade.qty;
      
      let revenue = (trade.type === "Crypto" && trade.totalSpent > 0) ? trade.totalSpent : (trade.qty * trade.unitPrice);
      s.totalSoldRevenue += revenue;
      
      s.cashFlows.push({ date: trade.date, amount: revenue }); 

      if (!portfolio.has(ticker)) continue;
      let activeTrades = portfolio.get(ticker);
      let sharesToSell = Number(trade.qty.toFixed(precision));
      let salePrice = trade.unitPrice;
      
      let sellPnL = 0; // Track PnL for this specific Sell execution

      // FIFO Loop to close lots (determines Unrealized basis)
      while (sharesToSell > 0 && activeTrades.length > 0) {
        let oldestTrade = activeTrades[0];
        let availableShares = Number(oldestTrade.shares.toFixed(precision));
        let sharesTaken = (availableShares <= sharesToSell) ? availableShares : sharesToSell;
        
        oldestTrade.shares = Number((oldestTrade.shares - sharesTaken).toFixed(precision));
        sharesToSell = Number((sharesToSell - sharesTaken).toFixed(precision));
        
        if (oldestTrade.shares === 0) activeTrades.shift();

        let costBasis = sharesTaken * oldestTrade.price;
        let revLot = sharesTaken * salePrice;
        
        let lotPnL = (revLot - costBasis);
        s.tradingPnL += lotPnL;
        sellPnL += lotPnL; 
      }
      
      // Update Trading Efficiency Metrics
      if (sellPnL > 0) {
          s.winningTrades += 1;
          s.grossProfit += sellPnL;
      } else if (sellPnL < 0) {
          s.losingTrades += 1;
          s.grossLoss += Math.abs(sellPnL);
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
    } else if (status === "CLOSED" && s.firstBuyDate && s.lastActionDate) {
      let timeDiff = s.lastActionDate.getTime() - s.firstBuyDate.getTime();
      daysHeld = Math.floor(timeDiff / (1000 * 3600 * 24));
    }

    let finalMinPrice = (s.minBuyPrice === Infinity) ? 0 : s.minBuyPrice;
    let avgSellPrice = (s.totalSoldShares > 0) ? (s.totalSoldRevenue / s.totalSoldShares) : 0;

    // --- APPLY REQUESTED OVERRIDE FOR CLOSED POSITIONS (Columns D & E) ---
    // If closed, show total sold shares as quantity, and average sell price as avg price
    let displayQty = (status === "CLOSED") ? s.totalSoldShares : tv.shares;
    let displayAvgPrice = (status === "CLOSED") ? avgSellPrice : avgPrice;

    let analysisNotes = [];
    if (tradingRoiPct < -0.01 && totalRoiPct > 0) analysisNotes.push("🐮 Cash Cow: Gain via Divs.");
    if (tradingRoiPct > 0.05 && (totalRoiPct - tradingRoiPct) < 0.01) analysisNotes.push("📈 Pure Trading.");
    if (finalMinPrice > 0 && finalMinPrice < (s.firstBuyPrice * 0.85)) analysisNotes.push("📉 Good DCA.");
    if (status === "OPEN" && avgPrice > (s.maxBuyPrice * 0.95) && s.tradeCount > 1) analysisNotes.push("⚠️ High Load Price.");
    if (status === "OPEN" && daysHeld > 730 && totalRoiPct < 0) analysisNotes.push("💤 Dead Money.");
    if (avgSellPrice > s.maxBuyPrice && avgSellPrice > 0) analysisNotes.push("💰 Sniper: Selling higher than buying.");

    let finalNote = analysisNotes.join(" | ");

    let currentMarketPrice = currentPrices.has(ticker) ? currentPrices.get(ticker) : avgPrice; 
    let finalMarketValue = tv.shares * currentMarketPrice;

    // --- 3A. UNREALIZED XIRR (Only open lots) ---
    let assetXirrUnrealized = 0;
    let noteUnrealized = status === "CLOSED" ? "Position closed. No Unrealized XIRR." : "Insufficient data.";
    
    if (status === "OPEN" && portfolio.has(ticker)) {
      let unrealizedCashFlows = [];
      let lots = portfolio.get(ticker);
      
      for (let lot of lots) {
        unrealizedCashFlows.push({ date: lot.date, amount: -(lot.shares * lot.price) });
      }
      if (tv.shares > 0.000001) {
        unrealizedCashFlows.push({ date: today, amount: finalMarketValue });
      }

      if (unrealizedCashFlows.length > 1) {
        assetXirrUnrealized = calculateInternalXIRR(unrealizedCashFlows);
        if (typeof assetXirrUnrealized === 'string') {
            noteUnrealized = assetXirrUnrealized;
            assetXirrUnrealized = 0;
        } else {
            let unrealizedRoiPct = tv.bookVal > 0 ? (finalMarketValue - tv.bookVal) / tv.bookVal : 0;
            let roiStr = (unrealizedRoiPct * 100).toFixed(1) + "%";
            let xirrStr = (assetXirrUnrealized * 100).toFixed(1) + "%";
            
            if (daysHeld > 0 && daysHeld < 365 && assetXirrUnrealized > unrealizedRoiPct) {
              noteUnrealized = `Short hold (${daysHeld} days): ${roiStr} Unrealized ROI accelerates XIRR to ${xirrStr}.`;
            } else if (daysHeld >= 365 && assetXirrUnrealized < unrealizedRoiPct) {
              noteUnrealized = `Long hold (${daysHeld} days): ${roiStr} Unrealized ROI diluted over years reduces XIRR to ${xirrStr}.`;
            } else {
              noteUnrealized = `Unrealized XIRR (${xirrStr}) naturally aligned with ROI (${roiStr}).`;
            }
        }
      }
    }

    // --- 3B. REALIZED XIRR (Lifetime History) ---
    let assetXirrRealized = 0;
    let noteRealized = "Insufficient data.";
    let realizedCashFlows = [...s.cashFlows]; 
    
    if (tv.shares > 0.000001) {
       realizedCashFlows.push({ date: today, amount: finalMarketValue });
    }

    if (realizedCashFlows.length > 1) {
      assetXirrRealized = calculateInternalXIRR(realizedCashFlows);
      if (typeof assetXirrRealized === 'string') {
          noteRealized = assetXirrRealized;
          assetXirrRealized = 0;
      } else {
          let roiStr = (totalRoiPct * 100).toFixed(1) + "%";
          let xirrStr = (assetXirrRealized * 100).toFixed(1) + "%";
          noteRealized = `Lifetime XIRR: ${xirrStr} vs Total Historical ROI: ${roiStr}.`;
      }
    }

    // --- 3C. NEW ADVANCED METRICS ---

    // 1. TWR (Modified Dietz) Annualized
    let dietzTWR = calculateModifiedDietz(s.cashFlows, finalMarketValue, s.firstBuyDate, today);
    
    // 2. Yield on Cost (Estimated Annualized)
    let yoc = 0;
    if (s.dividends > 0 && daysHeld > 0 && tv.bookVal > 0) {
        let annualDivs = s.dividends / (daysHeld / 365.25);
        yoc = annualDivs / tv.bookVal;
    }

    // 3. Trading Efficiency: Win Rate
    let totalClosedTrades = s.winningTrades + s.losingTrades;
    let winRate = totalClosedTrades > 0 ? (s.winningTrades / totalClosedTrades) : 0;

    // 4. Trading Efficiency: Profit Factor
    let profitFactor = s.grossLoss > 0 ? (s.grossProfit / s.grossLoss) : (s.grossProfit > 0 ? 999 : 0); 

    // 5. Drawdown
    let maxKnownPrice = Math.max(s.maxBuyPrice, s.maxSellPrice);
    let drawdown = (maxKnownPrice > 0 && currentMarketPrice < maxKnownPrice) 
                   ? (currentMarketPrice - maxKnownPrice) / maxKnownPrice 
                   : 0;

    // --- PUSH OUTPUT ---
    output.push([
      ticker, 
      s.type, 
      status, 
      displayQty,          // MODIFIED: Quantity (Column D) based on OPEN/CLOSED
      displayAvgPrice,     // MODIFIED: Avg Price (Column E) based on OPEN/CLOSED
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
      s.maxSellPrice, 
      avgSellPrice, 
      daysHeld,       
      s.lastActionDate, 
      s.tradeCount, 
      finalNote, 
      assetXirrUnrealized,  
      noteUnrealized,       
      assetXirrRealized,    
      noteRealized,         
      dietzTWR,             
      yoc,                  
      winRate,              
      profitFactor,         
      drawdown              
    ]);
  }
  return output;
}

/**
 * Helper: Calculates Annualized Modified Dietz Return (Proxy for TWR)
 * Evaluates the performance independent of timing and magnitude of cash flows.
 */
function calculateModifiedDietz(cashFlows, finalValue, startDate, endDate) {
  if (!startDate || !endDate || cashFlows.length === 0) return 0;
  
  let totalDays = (endDate.getTime() - startDate.getTime()) / (1000 * 3600 * 24);
  if (totalDays <= 0) return 0;

  let netCF = 0;
  let weightedCF = 0;

  for (let cf of cashFlows) {
    // Invert sign: Buys (negative in original cashFlows) become positive injections into the asset
    let amount = -cf.amount; 
    netCF += amount;

    let daysFromStart = (cf.date.getTime() - startDate.getTime()) / (1000 * 3600 * 24);
    let weight = (totalDays - daysFromStart) / totalDays;
    weight = Math.max(0, Math.min(1, weight)); // Safety cap

    weightedCF += (amount * weight);
  }

  let denominator = weightedCF;
  if (denominator === 0) return 0; // Avoid division by zero

  let dietzReturn = (finalValue - netCF) / denominator;
  
  let years = totalDays / 365.25;
  if (years <= 0) return 0;

  // Annualize the return
  if (dietzReturn > -1) {
     return Math.pow(1 + dietzReturn, 1 / years) - 1;
  } else {
     return -1; // Represents -100% loss
  }
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

    // 0. CAPTURE PREV ZAINETTO BEFORE ANY OPERATION
    let prevZainetto = 0;
    for (let amount of minusBucket.values()) prevZainetto += amount;

    // A. NET CURRENT YEAR GAINS AND LOSSES FIRST
    let netAnnualCompensable = tradingGain - s.lossAnno;
    
    let usedZainetto = 0;
    let taxableBase = gainNoComp; // Starts with Non-Compensable gains

    if (netAnnualCompensable > 0) {
      // We have a NET gain after applying this year's losses. 
      // Now use Zainetto (past losses) to offset the remaining gain.
      let availableYears = Array.from(minusBucket.keys()).sort((a, b) => a - b);
      let remainingGain = netAnnualCompensable;

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
      taxableBase += remainingGain;
    } else if (netAnnualCompensable < 0) {
      // We have a NET loss this year. Add it to the bucket.
      minusBucket.set(year, Math.abs(netAnnualCompensable));
    }

    // C. CLEAN EXPIRED LOSSES
    // FIX (1.14): si cattura l'importo in scadenza PRIMA di rimuoverlo dal basket.
    let aboutToExpire = minusBucket.has(expiryYearLimit) ? minusBucket.get(expiryYearLimit) : 0;
    if (minusBucket.has(expiryYearLimit)) {
      minusBucket.delete(expiryYearLimit);
    }

    // D. STRATEGIC ANALYSIS
    let totalResiduo = 0;
    let notes = [];
    let nextExpiryAmount = 0;

    // Per l'ANNO FISCALE CORRENTE le perdite in scadenza sono ancora recuperabili
    // (alert URGENT, "vendi azioni in gain"): non ancora perse -> expiredAmount = 0.
    // Per gli anni PASSATI sono effettivamente perse (alert EXPIRED).
    // Nota: i NUMERI fiscali (imponibile, tasse) non cambiano; cambia solo la nota consultiva.
    const isCurrentTaxYear = (year === new Date().getFullYear());
    if (isCurrentTaxYear) {
      nextExpiryAmount = aboutToExpire;
      expiredAmount = 0;
    } else {
      expiredAmount = aboutToExpire;
    }

    for (let [yOrigin, amount] of minusBucket.entries()) {
      totalResiduo += amount;
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
      s.lossAnno, 
      usedZainetto, expiredAmount, taxableBase, estimatedTax, totalResiduo, notes.join(" "), prevZainetto
    ]);
  }
  return report;
}


// --- 6. PUBLIC FUNCTIONS (For Google Sheet) ---

/**
 * PORTFOLIO DASHBOARD
 * Always shows everything (Stocks + Crypto)
 * @param {Array<Array>} currentPricesRange - OPTIONAL: Range with [Ticker, Current Market Price].
 * @customfunction
 */
function PORTFOLIO_DASHBOARD(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent, currentPricesRange) {
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true);
  
  let currentPrices = new Map();
  if (currentPricesRange && Array.isArray(currentPricesRange)) {
    for (let r = 0; r < currentPricesRange.length; r++) {
      let row = currentPricesRange[r];
      if (row && row[0] && row[1] !== undefined && row[1] !== "") {
        currentPrices.set(row[0].toString(), cleanNumber(row[1]));
      }
    }
  }

  return calculatePortfolioStats(stockData.concat(cryptoData), currentPrices);
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


// --- 9. ENGINE D: MONEY-WEIGHTED RETURN (XIRR) ---

/**
 * Calculates the REALIZED annualized Money-Weighted Return (XIRR).
 * Uses ONLY external account deposits and withdrawals where the Ticker is strictly "Cash".
 * @param {String} portfolioTarget - "ALL" (default), "STOCKS", or "CRYPTO" to calculate separately.
 * @customfunction
 */
function PORTFOLIO_XIRR(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent, currentPortfolioValue, evalDate, portfolioTarget = "ALL") {
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false, true);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true, true);
  
  let allData = [];
  let target = portfolioTarget ? portfolioTarget.toString().toUpperCase().trim() : "ALL";

  if (target === "STOCKS") {
    allData = stockData;
  } else if (target === "CRYPTO") {
    allData = cryptoData;
  } else {
    allData = stockData.concat(cryptoData); 
  }
  
  if (allData.length === 0) return "No data";
  
  let realizedCashFlows = [];
  
  for (let trade of allData) {
    if (!trade.date) continue;
    let tickerStr = trade.ticker.toString().toLowerCase().trim();
    let actionStr = trade.action.toString().toLowerCase().trim();
    
    let spent = parseFloat(trade.totalSpent);
    if (isNaN(spent) || spent === 0) {
        spent = Math.abs(parseFloat(trade.qty || 0)) * Math.abs(parseFloat(trade.price || 0));
    }
    
    // Strict constraint: Ticker must be "Cash"
    if (tickerStr === "cash") {
      if (actionStr === "deposit" || actionStr === "deposito") {
        realizedCashFlows.push({ date: trade.date, amount: -Math.abs(spent) });
      } else if (actionStr === "withdrawal" || actionStr === "prelievo") {
        realizedCashFlows.push({ date: trade.date, amount: Math.abs(spent) });
      }
    }
  }
  
  let finalDate = evalDate ? parseDate(evalDate) : new Date();
  let finalValue = cleanNumber(currentPortfolioValue);
  
  realizedCashFlows.push({ date: finalDate, amount: Math.abs(finalValue) });
  realizedCashFlows.sort((a, b) => a.date - b.date);
  
  return calculateInternalXIRR(realizedCashFlows);
}

/**
 * Calculates the UNREALIZED annualized Money-Weighted Return (XIRR).
 * Reconstructs open positions using FIFO logic and compares against current portfolio value.
 * @param {String} portfolioTarget - "ALL" (default), "STOCKS", or "CRYPTO" to calculate separately.
 * @customfunction
 */
function PORTFOLIO_XIRR_U(sDate, sSec, sAct, sQty, sPrice, sType, cDate, cSec, cAct, cQty, cPriceUnit, cSpent, currentPortfolioValue, evalDate, portfolioTarget = "ALL") {
  let stockData = normalizeData(sDate, sSec, sAct, sQty, sPrice, null, sType, false, true);
  let cryptoData = normalizeData(cDate, cSec, cAct, cQty, cPriceUnit, cSpent, null, true, true);
  
  let allData = [];
  let target = portfolioTarget ? portfolioTarget.toString().toUpperCase().trim() : "ALL";

  if (target === "STOCKS") {
    allData = stockData;
  } else if (target === "CRYPTO") {
    allData = cryptoData;
  } else {
    allData = stockData.concat(cryptoData); 
  }
  
  if (allData.length === 0) return "No data";
  
  let finalDate = evalDate ? parseDate(evalDate) : new Date();
  let finalValue = cleanNumber(currentPortfolioValue);
  
  let unrealizedCashFlows = [];
  let tradesByTicker = {};
  
  for (let trade of allData) {
    if (!trade.date) continue;
    let tStr = trade.ticker.toString().toUpperCase().trim();
    if (tStr === "CASH" || tStr === "") continue; 
    
    if (!tradesByTicker[tStr]) tradesByTicker[tStr] = [];
    tradesByTicker[tStr].push(trade);
  }
  
  for (let ticker in tradesByTicker) {
    let trades = tradesByTicker[ticker];
    trades.sort((a, b) => a.date - b.date); 
    
    let openLots = [];
    
    for (let trade of trades) {
      let actionStr = trade.action.toString().toLowerCase().trim();
      let qty = Math.abs(parseFloat(trade.qty) || 0);
      let spent = parseFloat(trade.totalSpent);
      if (isNaN(spent) || spent === 0) spent = qty * Math.abs(parseFloat(trade.price) || 0);
      
      if (qty === 0) continue;
      
      if (actionStr === "buy" || actionStr === "acquisto") {
        openLots.push({ date: trade.date, qty: qty, amount: spent });
      } else if (actionStr === "sell" || actionStr === "vendita") {
        let qtyToSell = qty;
        while (qtyToSell > 0 && openLots.length > 0) {
          let firstLot = openLots[0];
          if (firstLot.qty <= qtyToSell) {
            qtyToSell -= firstLot.qty;
            openLots.shift(); 
          } else {
            let proportion = qtyToSell / firstLot.qty;
            firstLot.qty -= qtyToSell;
            firstLot.amount -= firstLot.amount * proportion;
            qtyToSell = 0;
          }
        }
      }
    }
    
    for (let lot of openLots) {
      unrealizedCashFlows.push({ date: lot.date, amount: -Math.abs(lot.amount) });
    }
  }
  
  unrealizedCashFlows.push({ date: finalDate, amount: Math.abs(finalValue) });
  unrealizedCashFlows.sort((a, b) => a.date - b.date);
  
  return calculateInternalXIRR(unrealizedCashFlows);
}

/**
 * Helper: Newton-Raphson method for calculating Internal Rate of Return.
 * Upgraded with multi-guess for high volatility (Crypto) and validation.
 */
function calculateInternalXIRR(cashFlows) {
  if (cashFlows.length < 2) return "Error: Insufficient data (< 2 flows)";
  
  let hasPositive = false;
  let hasNegative = false;
  for (let cf of cashFlows) {
    if (cf.amount > 0) hasPositive = true;
    if (cf.amount < 0) hasNegative = true;
  }
  
  if (!hasNegative) return "Error: No investments or deposits found";
  if (!hasPositive) return "Error: Current value is zero or missing";

  const maxIterations = 100;
  const tolerance = 0.000001;
  let guesses = [0.1, -0.1, 0.5, -0.5, 0.0, -0.9, 1.0, 5.0]; 
  let startDate = cashFlows[0].date;
  
  for (let guess of guesses) {
    let rate = guess;
    
    for (let i = 0; i < maxIterations; i++) {
      let fValue = 0;
      let fDerivative = 0;
      
      for (let j = 0; j < cashFlows.length; j++) {
        let cf = cashFlows[j];
        let days = (cf.date - startDate) / (1000 * 60 * 60 * 24);
        let years = days / 365.0;
        
        fValue += cf.amount / Math.pow(1 + rate, years);
        fDerivative -= (years * cf.amount) / Math.pow(1 + rate, years + 1);
      }
      
      if (fDerivative === 0) break; 
      
      let newRate = rate - fValue / fDerivative;
      
      if (Math.abs(newRate - rate) < tolerance) {
        return newRate; 
      }
      rate = newRate;
    }
  }
  
  return "XIRR Error: Did not converge"; 
}
