/**
 * PHASE 1: Essential Data (UI Unblocker).
 * Loads critical metrics (Years, Dashboard, Budget, FIRE, Projections) immediately
 * to allow the user to interact with the interface while heavy data loads in the background.
 * * @param {number|string} year - The reference year.
 * @return {Object} Object containing all fast-loading data points.
 */
function loadFastStart(year) {
  const t1 = new Date().getTime();
  
  return {
    // Core Data
    years: getAvailableYears(),
    dashboard: getDashboardData(),
    savings: getMonthlySavings(),
    budget: getBudgetStatus(),
    transactions: getLastTransactions(),
    historyNW: getHistoryNWData(year),
    
    // --- NEW ADDITIONS (Immediate Calculation) ---
    // These were moved here to ensure the Dashboard looks complete instantly
    fire: getFireProgress(year),           // Calculates FIRE progress immediately
    projection: getYearlyProjection(year), // Calculates "Crystal Ball" immediately
    yearlyTotals: getYearlyTotals(year),   // Calculates Yearly Totals immediately
    runway: getRunwayData(year),           // Calculates Financial Runway immediately
    // ---------------------------------------------
    
    serverTime: new Date().getTime() - t1
  };
}

/**
 * PHASE 2: Heavy Content (Lazy Loading).
 * Loads data that requires complex calculation or external fetching (Charts, Portfolio).
 * This function is typically called asynchronously after the UI has rendered Phase 1.
 * * @param {number|string} year - The reference year.
 * @return {Object} Object containing heavy data points (Portfolio, Charts, etc.).
 */
function loadHeavyContent(year) {
  // NOTE: Watchlist removed to optimize performance
  return {
    portfolio: getLivePortfolio(), 
    pension: getPensionData(),
    charts: {
      financial: getMonthlyChartData(year),
      categories: getMonthlyCategoryBreakdown(year)
    },
    autocomplete: getAutocompleteData()
  };
}