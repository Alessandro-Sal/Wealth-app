/**
 * CUT-OFF DATE (Stocks Only):
 * - Stock positions closed AFTER this date will remain visible with 0 shares.
 * - Stock positions closed BEFORE this date will disappear.
 * - CRYPTO positions will ALWAYS disappear when closed, ignoring this date.
 */
const CUTOFF_DATE = new Date("2026-01-30"); 

/**
 * Creates a standardized Trade object.
 */
function createTradeObject(quantity, price, action) {
  return {
    shares: Number(quantity),
    price: Number(price),
    action: action
  };
}

/**
 * Core FIFO (First-In-First-Out) Calculation Engine.
 * 
 * @param {Array} security - List of tickers.
 * @param {Array} actions - List of actions (Buy/Sell).
 * @param {Array} quantity - List of quantities.
 * @param {Array} price - List of unit prices.
 * @param {Array} dates - List of trade dates (Essential for Stocks logic).
 * @param {number} precision - Decimal precision (5 for stocks, 8 for crypto).
 * @param {boolean} keepRows - If TRUE (Stocks), applies Cut-off logic. If FALSE (Crypto), always deletes closed rows.
 */
function processFifo(security, actions, quantity, price, dates, precision, keepRows) {
  let portfolio = new Map();

  for (let i = 0; i < security.length; i++) {
    let ticker = security[i].toString();
    if (!ticker) continue;
    
    let action = actions[i].toString();
    // Date handling: Mandatory for Stocks, optional but passed for Crypto
    let tradeDate = dates[i] ? new Date(dates[i]) : new Date("1900-01-01");
    
    let trade = createTradeObject(quantity[i], price[i], action);

    // --- BUY LOGIC ---
    if (action === "Buy" || action.toUpperCase() === "DRIP" || action.toUpperCase() === "REWARD") {
      if (!portfolio.has(ticker)) {
        portfolio.set(ticker, [trade]);
      } else {
        portfolio.get(ticker).push(trade);
      }
    }

    // --- SELL LOGIC (FIFO) ---
    if (action === "Sell" && portfolio.has(ticker)) {
      let activeTrades = portfolio.get(ticker);
      let sharesToSell = Number(trade.shares.toFixed(precision));

      while (sharesToSell > 0 && activeTrades.length > 0) {
        let oldestTrade = activeTrades[0];
        let availableShares = Number(oldestTrade.shares.toFixed(precision));

        if (availableShares <= sharesToSell) {
          // Fully consume the oldest lot
          sharesToSell = Number((sharesToSell - availableShares).toFixed(precision));
          activeTrades.shift(); 
        } else {
          // Partially reduce the oldest lot
          oldestTrade.shares = Number((availableShares - sharesToSell).toFixed(precision));
          sharesToSell = 0;
        }
      }

      // --- DELETION LOGIC (The "Zombie Row" Check) ---
      if (activeTrades.length === 0) {
        if (keepRows === true) {
          // --- STOCK LOGIC (Hybrid) ---
          // Only check date if it's a Stock.
          // If sold BEFORE cutoff: Delete. If sold AFTER: Keep row at 0.
          if (tradeDate < CUTOFF_DATE) {
             portfolio.delete(ticker);
          }
        } else {
          // --- CRYPTO LOGIC (Classic) ---
          // Always delete the row immediately when quantity hits 0.
          portfolio.delete(ticker);
        }
      }
    }
  }
  return portfolio;
}

/**
 * Interface for STOCKS (Sheet Function).
 * BEHAVIOR: Keeps empty rows if sold AFTER the cut-off date.
 * @customfunction
 */
function myStockPositions(security, actions, quantity, price, dates) {
  if (!security || !actions || !quantity || !price || !dates) return "Error: Select 5 columns (including Date)";
  // Pass 'true' to enable row retention logic
  const portfolio = processFifo(security, actions, quantity, price, dates, 5, true);
  return formatPortfolioOutput(portfolio);
}

/**
 * Interface for CRYPTO (Sheet Function).
 * BEHAVIOR: Always deletes the row when quantity is 0 (Classic behavior).
 * @customfunction
 */
function myCryptoPositions(security, actions, quantity, price, dates) {
  if (!security || !actions || !quantity || !price || !dates) return "Error: Select 5 columns (including Date)";
  // Pass 'false' to force deletion on close
  const portfolio = processFifo(security, actions, quantity, price, dates, 8, false);
  return formatPortfolioOutput(portfolio);
}

/**
 * Formats the Map into a 2D Array for Google Sheets.
 */
function formatPortfolioOutput(portfolioMap) {
  let results = [];
  
  portfolioMap.forEach((trades, ticker) => {
    let totalShares = 0;
    let totalCost = 0;
    
    trades.forEach(t => {
      totalShares += t.shares;
      totalCost += (t.shares * t.price);
    });

    let avgPrice = totalShares > 0 ? totalCost / totalShares : 0;
    results.push([ticker, totalShares, avgPrice]);
  });
  
  return results;
}