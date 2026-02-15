/**
 * ====================================================================
 * TAX OPTIMIZATION ENGINE (ITALIAN REGIME)
 * Handles Tax Loss Harvesting simulation and Capital Gains projections.
 * ====================================================================
 */

/**
 * Helper per pulire i numeri italiani (es. "1.250,50 €" -> 1250.50)
 */
function parseMyFloat(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  
  // Converte in stringa, rimuove €, spazi e punti delle migliaia
  let cleanStr = value.toString().replace(/[€\s]/g, "").replace(/\./g, "");
  // Sostituisce la virgola decimale con il punto
  cleanStr = cleanStr.replace(",", ".");
  
  let num = parseFloat(cleanStr);
  return isNaN(num) ? 0 : num;
}

/**
 * SCENARIO 1: TAX LOSS HARVESTING SCANNER
 * Scans open positions to identify "losers".
 * @customfunction
 */
function TAX_HARVESTING_SCANNER(tickers, quantities, avgPrices, currentPrices, assetType) {
  let report = [];
  report.push(["Ticker", "Qty", "Unrealized P&L", "Tax Credit (Minus)", "Note"]);

  const flatTickers = tickers ? tickers.map(r => r[0]) : [];
  const flatQty = quantities ? quantities.map(r => r[0]) : [];
  const flatAvg = avgPrices ? avgPrices.map(r => r[0]) : [];
  const flatCur = currentPrices ? currentPrices.map(r => r[0]) : [];
  
  let totalPotentialLoss = 0;

  for (let i = 0; i < flatTickers.length; i++) {
    let t = flatTickers[i];
    
    // Usa il parser sicuro sui numeri
    let q = parseMyFloat(flatQty[i]);
    let avg = parseMyFloat(flatAvg[i]);
    let cur = parseMyFloat(flatCur[i]);
    
    if (!t || q <= 0.00001) continue;
    
    let invested = q * avg;
    let currentVal = q * cur;
    let pnl = currentVal - invested;

    if (pnl < 0) {
      let potentialMinus = Math.abs(pnl);
      let note = "Sell to generate " + potentialMinus.toFixed(0) + "€ loss.";
      
      if (assetType === "ETF") {
         note += " (Valid for Basket)";
      }

      report.push([
        t,
        q,
        pnl.toFixed(2),
        potentialMinus.toFixed(2),
        note
      ]);
      
      totalPotentialLoss += potentialMinus;
    }
  }

  if (report.length === 1) {
    return [["No losing positions found.", "", "", "", ""]];
  }
  
  report.push(["TOTAL POTENTIAL", "", "", totalPotentialLoss.toFixed(2), "Total offset available"]);
  
  return report;
}

/**
 * SCENARIO 2: TAX BILL SIMULATOR
 * Calculates estimated tax.
 * @customfunction
 */
function ESTIMATE_TAX_BILL(realizedGainsYTD, availableZainetto, realizedLossesYTD) {
  // 1. Pulizia rigorosa degli input
  let gain = parseMyFloat(realizedGainsYTD);
  let zainetto = parseMyFloat(availableZainetto);
  let loss = parseMyFloat(realizedLossesYTD);

  let netPosition = gain - loss;
  
  let taxableBase = 0;
  let remainingZainetto = zainetto;
  let newZainettoGenerated = 0;

  if (netPosition > 0) {
    // Abbiamo un guadagno netto. Usiamo lo zainetto?
    if (remainingZainetto >= netPosition) {
      remainingZainetto -= netPosition;
      taxableBase = 0; // Compensato totalmente
    } else {
      taxableBase = netPosition - remainingZainetto;
      remainingZainetto = 0; // Finito tutto
    }
  } else {
    // Abbiamo una perdita netta quest'anno
    newZainettoGenerated = Math.abs(netPosition);
    taxableBase = 0;
  }

  let estimatedTax = taxableBase * 0.26;

  return [
    ["Metric", "Value", "Description"],
    ["Realized Gains (Compensable)", gain.toFixed(2), "Stock/Derivatives Gains"],
    ["Realized Losses (This Year)", loss.toFixed(2), "Losses closed in 2026"],
    ["Net Annual Position", netPosition.toFixed(2), "Gains - Losses"],
    ["Previous Basket (Zainetto)", zainetto.toFixed(2), "Losses from prev 4 years"],
    ["-----------------", "---", "---"],
    ["TAXABLE BASE", taxableBase.toFixed(2), "Amount subject to 26%"],
    ["ESTIMATED TAX DUE", estimatedTax.toFixed(2), "You pay this (approx)"],
    ["-----------------", "---", "---"],
    ["Residual Basket (Future)", remainingZainetto.toFixed(2), "Still available"],
    ["New Basket Generated", newZainettoGenerated.toFixed(2), "Added to Zainetto next year"]
  ];
}

