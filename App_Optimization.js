/**
 * PHASE 1: Essential Data (UI Unblocker).
 * Loads critical metrics (Years, Dashboard, Budget, FIRE, Projections) immediately
 * to allow the user to interact with the interface while heavy data loads in the background.
 * * @param {number|string} year - The reference year.
 * @return {Object} Object containing all fast-loading data points.
 */
function loadFastStart(year) {
  const currentYear = year || new Date().getFullYear().toString();
  return getFromCache('FAST_START_' + currentYear, () => {
    const t1 = new Date().getTime();
    return {
      years: getAvailableYears(),
      dashboard: getDashboardData(),
      savings: getMonthlySavings(),
      budget: getBudgetStatus(),
      transactions: getLastTransactions(),
      historyNW: getHistoryNWData(year),
      yearlyTotals: getYearlyTotals(year),
      balances: getBankBalances(),
      subsStatus: getSubsStatus(),
      credits: getActiveCredits(),
      serverTime: new Date().getTime() - t1
    };
  }, 1800); // 30 minutes cache
}

/**
 * PHASE 2: Heavy Content (Lazy Loading).
 * Loads data that requires complex calculation or external fetching (Charts, Portfolio).
 * This function is typically called asynchronously after the UI has rendered Phase 1.
 * * @param {number|string} year - The reference year.
 * @return {Object} Object containing heavy data points (Portfolio, Charts, etc.).
 */
function loadHeavyContent(year) {
  const currentYear = year || new Date().getFullYear().toString();
  return getFromCache('HEAVY_CONTENT_' + currentYear, () => {
    return {
      portfolio: getLivePortfolio(),
      pension: getPensionData(),
      charts: {
        financial: getMonthlyChartData(year),
        categories: getMonthlyCategoryBreakdown(year)
      },
      autocomplete: getAutocompleteData(),
      fire: getFireProgress(year),
      projection: getYearlyProjection(year),
      runway: getRunwayData(year)
    };
  }, 3600); // 1 hour cache
}

/**
 * FIX (2.12): refresh post-transazione in UN SOLO round-trip.
 * Il frontend (sendTransaction, ecc.) chiama oggi 5 funzioni separate
 * (fetchHistory + fetchBalances + fetchBudgets + fetchSubsStatus + fetchActiveCredits),
 * cioe' 5 chiamate google.script.run. Questa funzione le aggrega in una sola esecuzione
 * server -> 1 round-trip invece di 5 (post-transazione da ~15s a ~2-3s).
 *
 * WIRING FRONTEND (da fare manualmente, vedi README MODIFICHE_FASE2):
 *   google.script.run.withSuccessHandler(d => {
 *     renderLast5(d.history); renderBalances(d.balances); renderBudgets(d.budgets);
 *     renderSubs(d.subsStatus); renderCredits(d.credits);
 *   }).getPostTransactionRefresh();
 *
 * @return {Object} { history, balances, budgets, subsStatus, credits }
 */
function getPostTransactionRefresh() {
  return {
    history: getLastTransactions(),
    balances: getBankBalances(),
    budgets: getBudgetStatus(),
    subsStatus: getSubsStatus(),
    credits: getActiveCredits()
  };
}