/**
 * Scans all sheets in the Spreadsheet to find those named "Expenses Tracker [Year]".
 * Extracts the years and returns them as a list, sorted from newest to oldest.
 * * @return {Array<string>} Array of years (e.g., ["2026", "2025", "2024"]).
 */
function getExpenseSheetYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let years = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    // Check if name starts with "Expenses Tracker " (note the trailing space)
    if (name.startsWith("Expenses Tracker ")) {
      // Extract the year part
      let year = name.replace("Expenses Tracker ", "").trim();
      // If it's a valid number, add to list
      if (!isNaN(parseFloat(year))) {
        years.push(year);
      }
    }
  });

  // Sort descending (Newest first)
  return years.sort().reverse();
}

/**
 * Scans "History" sheets (Stocks/Crypto) to find all years where transactions occurred.
 * useful for populating dropdown menus in the UI.
 * * @return {Array<number>} Unique array of years sorted descending.
 */
function getInvestmentYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["History B/S Stocks", "History B/S Crypto"];
  let yearsSet = new Set();

  // Always add current year for convenience
  yearsSet.add(new Date().getFullYear());

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() >= 3) {
      // Read only Column A (Date) starting from Row 3
      const dates = sheet.getRange(3, 1, sheet.getLastRow() - 2, 1).getValues();
      dates.forEach(row => {
        if (row[0] instanceof Date) {
          yearsSet.add(row[0].getFullYear());
        }
      });
    }
  });

  // Convert Set to Array and sort descending
  return Array.from(yearsSet).sort((a, b) => b - a);
}

/**
 * --- YEARLY TOTALS (RECAP) ---
 * Calculates Total Income, Expenses, and Savings for a specific year.
 * Aggregates data by category and formats the output for the UI.
 * * @param {string|number} year - The specific year to analyze.
 * @return {Object} Object containing formatted totals and sorted category lists.
 */
function getYearlyTotals(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Year and Target Sheet
  let targetYear = String(year).trim();
  if (!targetYear || targetYear === "undefined") targetYear = String(new Date().getFullYear());
  
  const sheetName = "Expenses Tracker " + targetYear;
  const sheet = ss.getSheetByName(sheetName);

  let totalIncomeVal = 0;
  let totalExpenseVal = 0;
  
  // Maps to accumulate category totals
  let incomeMap = {};
  let expenseMap = {};

  // Default / Empty values
  const fmt = (n) => "€ " + n.toLocaleString("it-IT", {minimumFractionDigits: 0, maximumFractionDigits: 0});
  const emptyRes = { income: "€ 0", expense: "€ 0", saving: "€ 0", rawSaving: 0, incCats: [], expCats: [] };

  if (!sheet || sheet.getLastRow() < 20) return emptyRes;

  // 2. Read Data (From Row 20 downwards)
  const data = sheet.getRange(20, 1, sheet.getLastRow() - 19, 10).getValues();

  data.forEach(row => {
    if (!row[0] || !row[1]) return;
    
    let date = new Date(row[0]);
    
    // Year Filter
    if (!isNaN(date) && date.getFullYear() == targetYear) {
      let type = String(row[1]);
      let cat = String(row[2]).trim(); // Category is in Column C
      
      let amount = 0;
      // Sum Bank Columns (Indices 4 to 9 / Cols E->J)
      for (let i = 4; i <= 9; i++) {
        let val = parseFloat(row[i]);
        if (!isNaN(val)) amount += val;
      }
      
      if (amount === 0) return;

      // 3. Accumulate Totals and Map Categories
      if (type === 'Expense') {
        amount = Math.abs(amount);
        totalExpenseVal += amount;
        expenseMap[cat] = (expenseMap[cat] || 0) + amount;
      } 
      else if (type === 'Income') {
        totalIncomeVal += amount;
        incomeMap[cat] = (incomeMap[cat] || 0) + amount;
      }
    }
  });

  let netSaving = totalIncomeVal - totalExpenseVal;

  // 4. Helper to sort categories by value (Descending)
  const sortCats = (map) => {
    return Object.keys(map).map(key => {
      return { cat: key, val: map[key] };
    }).sort((a,b) => b.val - a.val); 
  };

  return {
    income: fmt(totalIncomeVal),
    expense: fmt(totalExpenseVal),
    saving: fmt(netSaving),
    rawSaving: netSaving,
    incCats: sortCats(incomeMap), // Sorted Income Categories
    expCats: sortCats(expenseMap) // Sorted Expense Categories
  };
}

/**
 * --- HISTORICAL DATA (Analytical Net Worth) ---
 * Retrieves historical snapshot data from the "NW analitico" sheet.
 * Maps specific rows (Liquid, Stocks, ETFs, Crypto, Pension) based on the requested year column.
 * * @param {string|number} year - The year column to retrieve.
 * @return {Object} Structured historical data for charts or reports.
 */
function getHistoryNWData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("NW analitico");
  
  if (!sheet) return { error: "Sheet 'NW analitico' not found" };

  const data = sheet.getDataRange().getDisplayValues(); 
  const headers = data[0]; 

  // --- 1. SETUP HELPERS ---
  const parseVal = (val) => {
    if (!val) return 0;
    let s = String(val).replace(/[^0-9.,-]/g, '');
    // Handle European vs US number formats
    if (s.includes(',') && s.includes('.')) {
       s = s.replace(/\./g, '').replace(',', '.'); 
    } else if (s.includes(',')) {
       s = s.replace(',', '.');
    }
    let num = parseFloat(s);
    return isNaN(num) ? 0 : num;
  };

  // Find the column index for the requested year
  let colIndex = headers.indexOf(String(year));
  if (colIndex === -1) {
    colIndex = headers.findIndex(h => String(h).includes(String(year)));
  }
  // Fallback: If year not found, use the last available column
  if (colIndex === -1) colIndex = headers.length - 1;

  // Helper to get value from a specific row (1-based row number)
  const getRowVal = (rowNum) => {
    const rIndex = rowNum - 1; 
    if (rIndex < 0 || rIndex >= data.length) return 0;
    return parseVal(data[rIndex][colIndex]);
  };

  // Helper to get percentage string
  const getRowPctStr = (rowNum) => {
    const rIndex = rowNum - 1;
    if (rIndex < 0 || rIndex >= data.length) return "0.00%";
    let val = data[rIndex][colIndex];
    if (!val) return "0.00%";
    if (String(val).includes('%')) return val;
    return (parseFloat(String(val).replace(',', '.')) * 100).toFixed(2) + "%";
  };

  // --- 2. READ DATA FROM "NW ANALITICO" ---
  
  // Totals
  const valLiquidEur = getRowVal(2); // Liquid NW
  const valLiquidUsd = getRowVal(4); 
  const valTotalUsd  = getRowVal(8); 

  // Asset Classes
  const vStocks  = getRowVal(44);
  const vEtfs    = getRowVal(35);
  const vCrypto  = getRowVal(16);
  const vCash    = getRowVal(10);
  const vCashEq  = getRowVal(12);
  const vOthers  = getRowVal(14);
  
  // --- UPDATE: PENSION FROM ROW 76 ---
  let vPension = getRowVal(76); 

  // --- 3. RECALCULATE TOTALS AND ALLOCATION ---
  
  // Total Liquidity logic: Use CashEq if available, otherwise Cash
  const totalLiquidity = vCashEq > 0 ? vCashEq : vCash;

  // Real Total = Sum of all assets (including Pension Row 76)
  const calculatedTotal = vStocks + vEtfs + vCrypto + vOthers + vPension + totalLiquidity;
  
  const calcAlloc = (val) => {
    if (!calculatedTotal || calculatedTotal === 0) return "0.00%";
    return ((val / calculatedTotal) * 100).toFixed(2) + "%";
  };

  return {
    year: year,
    liquid: { val: valLiquidEur, usd: valLiquidUsd },
    total:  { val: calculatedTotal, usd: valTotalUsd }, // Recalculated total for consistency
    
    summary: {
      stocks:  { val: vStocks,  pct: calcAlloc(vStocks) },
      cash:    { val: vCash,    pct: calcAlloc(vCash) },
      cashEq:  { val: vCashEq,  pct: calcAlloc(vCashEq) },
      etfs:    { val: vEtfs,    pct: calcAlloc(vEtfs) },
      crypto:  { val: vCrypto,  pct: calcAlloc(vCrypto) },
      others:  { val: vOthers,  pct: calcAlloc(vOthers) },
      pension: { val: vPension, pct: calcAlloc(vPension) } 
    },
    
    // Specific Section Details (Unchanged logic)
    details: {
      crypto: {
        main: getRowVal(16), invested: getRowVal(17),
        balVal: getRowVal(18), balPct: getRowPctStr(21),
        unrVal: getRowVal(19), unrPct: getRowPctStr(22),
        realVal: getRowVal(20), realPct: getRowPctStr(23)
      },
      stocks: {
        main: getRowVal(44), invested: getRowVal(45),
        balVal: getRowVal(46), balPct: getRowPctStr(49),
        unrVal: getRowVal(47), unrPct: getRowPctStr(50),
        realVal: getRowVal(48), realPct: getRowPctStr(51)
      },
      etfs: {
        main: getRowVal(35), invested: getRowVal(36),
        balVal: getRowVal(37), balPct: getRowPctStr(40),
        unrVal: getRowVal(38), unrPct: getRowPctStr(41),
        realVal: getRowVal(39), realPct: getRowPctStr(42)
      }
    }
  };
}

/**
 * Retrieves historical trend data for Net Worth and Asset Classes over time.
 * Extracts data from "NW analitico" strictly for columns where the header is an exact year.
 * * @return {Object} Trend data containing labels (years) and datasets for each asset class, including Cash, Cash Eq, and USD equivalents.
 */
function getHistoricalNetWorthTrend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("NW analitico");
  
  if (!sheet) return { error: "Sheet 'NW analitico' not found" };

  const data = sheet.getDataRange().getValues();
  const headers = data[0]; 

  // Helper to safely parse string values into numbers
  const parseVal = (val) => {
    if (!val) return 0;
    let s = String(val).replace(/[^0-9.,-]/g, '');
    if (s.includes(',') && s.includes('.')) {
       s = s.replace(/\./g, '').replace(',', '.'); 
    } else if (s.includes(',')) {
       s = s.replace(',', '.');
    }
    let num = parseFloat(s);
    return isNaN(num) ? 0 : num;
  };

  let result = {
    labels: [],
    totalNW: [],
    liquidNW: [],
    totalNW_USD: [], // Added Total USD
    liquidNW_USD: [], // Added Liquid USD
    stocks: [],
    etfs: [],
    crypto: [],
    pension: [],
    cash: [],
    cashEq: []
  };

  // Loop through columns starting from index 1 (skipping row labels in column A)
  for (let c = 1; c < headers.length; c++) {
    let rawLabel = headers[c];
    if (!rawLabel) continue; 
    
    let labelStr = String(rawLabel).trim();

    // ONLY process columns where the header is exactly a 4-digit year
    if (!/^\d{4}$/.test(labelStr)) {
      continue;
    }
    
    result.labels.push(labelStr);

    // Retrieve values mapping the correct rows
    let valLiquidUsd = parseVal(data[3][c]); // Row 4 -> index 3
    let valTotalUsd  = parseVal(data[7][c]); // Row 8 -> index 7

    let vStocks = parseVal(data[43][c]);     // Row 44 -> index 43
    let vEtfs = parseVal(data[34][c]);       // Row 35 -> index 34
    let vCrypto = parseVal(data[15][c]);     // Row 16 -> index 15
    let vCash = parseVal(data[9][c]);        // Row 10 -> index 9
    let vCashEq = parseVal(data[11][c]);     // Row 12 -> index 11
    let vPension = parseVal(data[75][c]);    // Row 76 -> index 75

    // 1) Liquid NW Calculation = Cash + Stocks + Crypto + Etfs
    let calcLiquidNW = vCash + vStocks + vCrypto + vEtfs;

    // 2) Total NW Calculation = Cash + Cash Eq + Stocks + Crypto + Etfs + Pension
    let calcTotalNW = vCashEq + vStocks + vCrypto + vEtfs + vPension;

    result.totalNW.push(calcTotalNW);
    result.liquidNW.push(calcLiquidNW);
    result.totalNW_USD.push(valTotalUsd);
    result.liquidNW_USD.push(valLiquidUsd);
    result.stocks.push(vStocks);
    result.etfs.push(vEtfs);
    result.crypto.push(vCrypto);
    result.pension.push(vPension);
    result.cash.push(vCash);
    result.cashEq.push(vCashEq);
  }

  return result;
}