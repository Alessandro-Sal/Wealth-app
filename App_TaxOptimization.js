/**
 * Retrieves Tax Optimization data including estimated tax bill, loss harvesting potential,
 * and expiring tax credits (Zainetto) by processing raw transaction data.
 * @return {Object} Structured JSON data for the UI.
 */
function getTaxOptimizationData() {
  try {
    // Helper to parse Italian/European formatted numbers from DisplayValues
    function _parseLocNumber(str) {
      if (typeof str === "number") return str;
      if (!str) return 0;
      let isNegative = String(str).includes('-') || (String(str).includes('(') && String(str).includes(')'));
      let s = String(str).replace(/[^\d.,]/g, '').trim();
      if (!s) return 0;
      let lastComma = s.lastIndexOf(',');
      let lastDot = s.lastIndexOf('.');
      if (lastComma > lastDot) {
        s = s.replace(/\./g, '').replace(',', '.');
      } else if (lastDot > lastComma) {
        s = s.replace(/,/g, '');
      } else if (lastComma !== -1) {
        s = s.replace(',', '.');
      }
      let val = parseFloat(s) || 0;
      return isNegative ? -val : val;
    }

    // 1. Fetch raw transactions
    let allData = getAllTransactions();
    
    // 2. Calculate Zainetto
    let zainettoReport = calculateZainettoReport(allData);
    
    // 3. Extract Current Year Data
    let currentYear = new Date().getFullYear();
    let currentYearData = zainettoReport.find(r => r[0] === currentYear);
    
    // Fallback if no activity in current year
    if (!currentYearData) {
      let lastRow = zainettoReport[zainettoReport.length - 1];
      let residualBasket = lastRow ? lastRow[8] : 0;
      currentYearData = [currentYear, 0, 0, 0, 0, 0, 0, 0, residualBasket, ""];
    }
    
    let taxBillData = [];
    let isExpiringThisYear = false;
    let expiringAmount = 0;
    
    if (currentYearData) {
      let gainCompensabile = currentYearData[1]; // Compensable Gain
      let gainNonComp = currentYearData[2];      // Non Compensable Gain
      let newLosses = currentYearData[3];        // Losses from this year
      let usedZainetto = currentYearData[4];     // Zainetto consumed
      let expiredAmount = currentYearData[5];    // Expired
      let taxableBase = currentYearData[6];      // Net taxable base
      let estimatedTax = currentYearData[7];     // Tax due
      let residualBasket = currentYearData[8];   // Zainetto remaining for next year
      let notes = currentYearData[9] || "";
      let prevZainetto = currentYearData[10] || 0; // Zainetto imported from last year

      // Tax Bill
      taxBillData.push({ metric: "Realized Gains (Compensable)", value: gainCompensabile, desc: "Stock/Derivatives Gains" });
      taxBillData.push({ metric: "Realized Gains (Non-Compensable)", value: gainNonComp, desc: "ETF/Dividends Gains" });
      taxBillData.push({ metric: "Realized Losses (This Year)", value: newLosses, desc: "Losses closed in " + currentYear });
      taxBillData.push({ metric: "Previous Basket (Zainetto)", value: prevZainetto, desc: "Losses from prev 4 years" });
      taxBillData.push({ metric: "TAXABLE BASE", value: taxableBase, desc: "Amount subject to 26%" });
      taxBillData.push({ metric: "ESTIMATED TAX DUE", value: estimatedTax, desc: "You pay this (approx)" });
      taxBillData.push({ metric: "Residual Basket (Future)", value: residualBasket, desc: "Still available" });
      
      if (notes.includes("URGENT:")) {
        isExpiringThisYear = true;
        let match = notes.match(/URGENT: (\d+)€/);
        if (match) expiringAmount = parseInt(match[1]);
      }
    }
    
    // 4. Calculate Harvesting Opportunities using Live Portfolio
    let livePort = null;
    try {
      livePort = getLivePortfolio();
    } catch(e) {
      console.warn("Could not get live portfolio for harvesting: " + e);
    }
    
    let harvestingData = [];
    let totalPotential = 0;
    
    if (livePort) {
      let assets = [];
      if (livePort.stocks) assets = assets.concat(livePort.stocks);
      if (livePort.etfs) assets = assets.concat(livePort.etfs);
      if (livePort.crypto && INCLUDE_CRYPTO_IN_TAX_BASKET) assets = assets.concat(livePort.crypto);
      
      assets.forEach(asset => {
        let t = asset.t;
        let q = _parseLocNumber(asset.qty);
        let avg = _parseLocNumber(asset.avgPrice);
        let cur = _parseLocNumber(asset.price);
        let type = asset.sector === 'ETFs' ? "ETF" : "Stock"; 
        
        if (!t || q <= 0.00001) return;
        
        let invested = q * avg;
        let currentVal = q * cur;
        let pnl = currentVal - invested;
    
        if (pnl < 0) {
          let potentialMinus = Math.abs(pnl);
          let note = "Sell to generate " + potentialMinus.toFixed(0) + "€ loss.";
          if (type === "ETF") note += " (Valid for Basket)";
          
          harvestingData.push({
            ticker: t,
            qty: q,
            unrealized: pnl,
            minus: potentialMinus,
            note: note
          });
          totalPotential += potentialMinus;
        }
      });
    }

    return {
      bill: taxBillData,
      harvesting: harvestingData,
      totalPotential: totalPotential,
      zainettoDetails: {
        isExpiringThisYear: isExpiringThisYear,
        expiringAmount: expiringAmount
      }
    };
  } catch (error) {
    console.error("Error in getTaxOptimizationData: " + error);
    return { error: error.toString() };
  }
}

/**
 * Retrieves all transactions from DB sheets, normalized.
 */
function getAllTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("History B/S Stocks");
  const cryptoSheet = ss.getSheetByName("History B/S Crypto");
  let allData = [];

  if (stockSheet && stockSheet.getLastRow() >= 3) {
    const data = stockSheet.getRange(3, 1, stockSheet.getLastRow() - 2, 8).getValues();
    data.forEach(r => {
      let date = r[0] ? new Date(r[0]) : new Date();
      let ticker = String(r[1]).trim();
      let action = String(r[2]).trim();
      let qty = parseFloat(r[3]) || 0;
      let type = String(r[4]).trim() || "Stock";
      let unitPrice = parseFloat(r[5]) || 0;
      let totalSpent = parseFloat(r[7]) || (qty * unitPrice); 
      if(action === "Dividend" || action === "Dividends") {
          totalSpent = unitPrice;
          qty = 0;
      }
      if (ticker && !ticker.toUpperCase().includes("CASH") && !ticker.toUpperCase().includes("EUR")) {
        allData.push({ date, ticker, action, qty, unitPrice, totalSpent, type });
      }
    });
  }

  // INCLUDE_CRYPTO_IN_TAX_BASKET is defined globally in Sheets_Taxes & RealGains.js
  let includeCrypto = false;
  try { includeCrypto = INCLUDE_CRYPTO_IN_TAX_BASKET; } catch(e) {}

  if (cryptoSheet && cryptoSheet.getLastRow() >= 3 && includeCrypto) {
    const data = cryptoSheet.getRange(3, 1, cryptoSheet.getLastRow() - 2, 8).getValues();
    data.forEach(r => {
      let date = r[0] ? new Date(r[0]) : new Date();
      let ticker = String(r[1]).trim();
      let action = String(r[2]).trim();
      let qty = parseFloat(r[3]) || 0;
      let type = "Crypto";
      let totalSpent = parseFloat(r[4]) || 0; // Col E
      let unitPrice = parseFloat(r[5]) || (qty !== 0 ? totalSpent / qty : 0);
      
      if (ticker && !ticker.toUpperCase().includes("CASH") && !ticker.toUpperCase().includes("EUR")) {
        allData.push({ date, ticker, action, qty, unitPrice, totalSpent, type });
      }
    });
  }

  allData.sort((a, b) => a.date - b.date);
  return allData;
}

/**
 * WHAT-IF SIMULATOR
 * Simulates a fake sale and calculates the impact on the Zainetto and Taxable Base.
 */
function simulateSale(ticker, qty, simulatedPrice) {
  try {
    let allData = getAllTransactions();
    let cleanTicker = String(ticker).trim().toUpperCase();
    
    let existingLots = allData.filter(d => d.ticker === cleanTicker);
    let typeToUse = existingLots.length > 0 ? existingLots[0].type : "Stock";
    let currentYear = new Date().getFullYear();

    // 1. Get Current State
    let currentZainetto = calculateZainettoReport(allData);
    let currentYearData = currentZainetto.find(r => r[0] === currentYear) || [currentYear, 0, 0, 0, 0, 0, 0, 0, (currentZainetto[currentZainetto.length - 1] ? currentZainetto[currentZainetto.length - 1][8] : 0), ""];
    
    let oldTaxableBase = currentYearData[6];
    let oldTaxDue = currentYearData[7];
    let oldResidualBasket = currentYearData[8];

    // 2. Inject the fake transaction TODAY
    allData.push({
      date: new Date(),
      ticker: cleanTicker,
      action: "Sell",
      qty: parseFloat(qty),
      unitPrice: parseFloat(simulatedPrice),
      totalSpent: parseFloat(qty) * parseFloat(simulatedPrice),
      type: typeToUse
    });
    
    // 3. Get Simulated State
    let simulatedZainetto = calculateZainettoReport(allData);
    let simulatedYearData = simulatedZainetto.find(r => r[0] === currentYear);
    
    if (!simulatedYearData) return { error: "Could not calculate simulated year." };
    
    let newTaxableBase = simulatedYearData[6];
    let newTaxDue = simulatedYearData[7];
    let newResidualBasket = simulatedYearData[8];

    return {
      success: true,
      simulatedTaxableBase: newTaxableBase,
      simulatedTaxDue: newTaxDue,
      simulatedResidualBasket: newResidualBasket,
      deltaTaxableBase: newTaxableBase - oldTaxableBase,
      deltaTaxDue: newTaxDue - oldTaxDue,
      deltaResidualBasket: newResidualBasket - oldResidualBasket,
      inferredType: typeToUse
    };
  } catch (error) {
    return { error: error.toString() };
  }
}