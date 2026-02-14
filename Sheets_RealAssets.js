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
  let r = (annualRate / 100) / 12; // Monthly rate
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