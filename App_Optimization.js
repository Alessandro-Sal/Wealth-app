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
    // Core Data (servono per il primo paint istantaneo della dashboard)
    years: getAvailableYears(),
    dashboard: getDashboardData(),
    savings: getMonthlySavings(),
    budget: getBudgetStatus(),
    transactions: getLastTransactions(),
    historyNW: getHistoryNWData(year),
    yearlyTotals: getYearlyTotals(year),

    // FIX (1.18): fire / projection / runway SPOSTATE in loadHeavyContent.
    // Erano tra i calcoli piu' pesanti (leggono "Expenses Tracker" intero) e gonfiavano
    // il primo round-trip a ~5-8s. renderAllData() usa guardie (if data.fire / .projection /
    // .runway) ed e' richiamata anche sui dati "heavy", quindi compariranno subito dopo,
    // senza errori, mentre la dashboard principale e' gia' interattiva.

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
    autocomplete: getAutocompleteData(),

    // FIX (1.18): spostate qui da loadFastStart (vedi sopra). renderAllData() le rende
    // appena arriva questo payload, tramite le sue guardie if(data.fire/.projection/.runway).
    fire: getFireProgress(year),
    projection: getYearlyProjection(year),
    runway: getRunwayData(year)
  };
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