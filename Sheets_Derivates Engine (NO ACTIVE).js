/**
 * ====================================================================
 * DERIVATIVES ENGINE (OPTIONS & FUTURES)
 * Manages Short Selling, Multipliers, and P&L
 * ====================================================================
 */

function calculateDerivatives(tickers, actions, quantities, prices, multipliers, types) {
  let report = [];
  let inventory = new Map();

  // 1. Input Normalization (Raw Data)
  const flatTickers = tickers ? tickers.map(r => r[0]) : [];
  const flatActions = actions ? actions.map(r => r[0]) : [];
  const flatQtys = quantities ? quantities.map(r => r[0]) : [];
  const flatPrices = prices ? prices.map(r => r[0]) : [];
  const flatMults = multipliers ? multipliers.map(r => r[0]) : []; // Column for Multiplier (e.g., 100 for options)
  const flatTypes = types ? types.map(r => r[0]) : []; // Call, Put, Future

  // 2. Chronological Processing
  // 
  for (let i = 0; i < flatTickers.length; i++) {
    let ticker = flatTickers[i];
    if (!ticker || ticker === "") continue;

    let action = flatActions[i].toString().toUpperCase(); // BUY, SELL, EXPIRE, ASSIGN
    let qty = Number(flatQtys[i]); 
    let price = Number(flatPrices[i]);
    let mult = flatMults[i] ? Number(flatMults[i]) : 100; // Default 100 for options if empty
    let type = flatTypes[i] ? flatTypes[i].toString() : "Option";

    if (!inventory.has(ticker)) {
      inventory.set(ticker, {
        netQty: 0,
        totalCost: 0,
        realizedPnL: 0,
        type: type,
        multiplier: mult,
        history: [] // For debug or detailed analysis
      });
    }

    let pos = inventory.get(ticker);
    
    // Cashflow Calculation
    // Buy = Money out (Negative), Sell = Money in (Positive)
    let cashflow = 0;
    
    if (action === "BUY" || action === "BTC" || action === "BUY TO CLOSE") {
      // Purchase (Long Entry or Short Closing)
      cashflow = -(qty * price * mult);
      pos.netQty += qty;
      pos.totalCost += (-cashflow); // Add to "cost spent"
    } 
    else if (action === "SELL" || action === "STO" || action === "SELL TO OPEN") {
      // Sale (Short Entry or Long Closing)
      cashflow = (qty * price * mult);
      pos.netQty -= qty;
      pos.realizedPnL += cashflow; // Temporarily add all to cashflow, balance later
    }
    else if (action === "EXPIRE") {
      // Expiration at zero value
      // If Long (Qty > 0): Lost the entire cost.
      // If Short (Qty < 0): Kept the entire premium collected.
      pos.netQty = 0; // Reset position
      // No cashflow occurs at expiration (if OTM), it is just an accounting closure
    }
    
    // Simplified P&L Logic for Derivatives (Average Cost)
    // Note: Derivatives are complex. This calculates P&L based on "Net at end of day".
    // If the position is CLOSED (netQty == 0), the entire historical net cashflow is Realized P&L.
  }

  // 3. Output Generation
  // Format: Ticker | Type | Status | Open Contracts | Avg Price | Realized P&L | Notional Exposure
  
  let resultKeys = Array.from(inventory.keys()).sort();
  
  for (let t of resultKeys) {
    let pos = inventory.get(t);
    let status = (Math.abs(pos.netQty) > 0.001) ? "OPEN" : "CLOSED";
    
    // Intelligent Realized P&L Calculation
    // P&L = (Sum of all inflows) - (Sum of all outflows)
    // Warning: If still OPEN, part of the cost is "Invested", not "Lost".
    
    // Here we use a simplified approach for derivatives: 
    // We track every trade. But for a simple report, it is often enough to know:
    // "How much I spent vs How much I collected" on that specific ticker.
    
    // We reconstruct precise P&L by iterating movements (Average Cost method for simplicity on derivatives)
    // Note: For perfection, a bidirectional FIFO (First-In-First-Out) is needed (complex), here we use the Cashflow method.
    
    // Average Price Calculation (Break Even)
    let avgPrice = 0;
    if (pos.netQty !== 0) {
      // If Long, what was the average pay?
      // If Short, what was the average sell price?
      // This would require saving lots.
      // For now, leave empty or approximated.
    }

    report.push([
      t,                // Unique Ticker
      pos.type,         // Call/Put/Future
      status,           // Open/Closed
      pos.netQty,       // How many contracts held (+ Long, - Short)
      pos.multiplier,   // Multiplier used
      // Here you should calculate P&L. 
      // If CLOSED: Total Cashflow. 
      // If OPEN: Do not calculate P&L until closed, or show mark-to-market if current price is passed.
      "See Notes"      
    ]);
  }
  
  return report;
}

/**
 * Calculates derivative positions for Sheets.
 * @param tickers Ticker Column Range (e.g., AAPL_240101_C_150)
 * @param actions Actions Column Range (Buy/Sell/Expire)
 * @param qtys Quantities Column Range
 * @param prices Unit Price Column Range
 * @param mults Multiplier Column Range (e.g., 100)
 * @customfunction
 */
function REPORT_DERIVATI(tickers, actions, qtys, prices, mults) {
  // Simplified logic based on total cash flows per Unique Ticker
  // 
  let map = new Map();

  // Array safe check
  let T = tickers || [];
  let A = actions || [];
  let Q = qtys || [];
  let P = prices || [];
  let M = mults || [];

  for (let i = 0; i < T.length; i++) {
    let name = T[i][0];
    if (!name) continue;
    let action = A[i][0].toString().toUpperCase();
    let q = Number(Q[i][0]);
    let p = Number(P[i][0]);
    let m = M[i] && M[i][0] ? Number(M[i][0]) : 100; // Default 100

    if (!map.has(name)) map.set(name, { qty: 0, cashflow: 0, openCost: 0 });
    let obj = map.get(name);

    let flow = q * p * m;

    if (action.includes("BUY")) {
      obj.qty += q;           // Increase contracts (or reduce negative)
      obj.cashflow -= flow;   // Money out
    } else if (action.includes("SELL")) {
      obj.qty -= q;           // Reduce contracts (or go negative/short)
      obj.cashflow += flow;   // Money in
    } else if (action.includes("EXPIRE")) {
      // If it expires, reset quantity. Cashflow remains historical (profit if short, loss if long)
      obj.qty = 0;
    }
  }

  let output = [["Contract", "Status", "Net Position", "Realized P&L (Cashflow)", "Note"]];
  
  for (let [key, val] of map) {
    let status = (val.qty === 0) ? "CLOSED" : "OPEN";
    
    // If CLOSED, the residual Cashflow is the Net P&L.
    // If OPEN, the Cashflow is partial (does not include current closing value).
    let pnl = val.cashflow; 
    let note = (status === "OPEN") ? "Provisional P&L (excludes current value)" : "Final P&L";

    output.push([key, status, val.qty, pnl, note]);
  }
  return output;
}