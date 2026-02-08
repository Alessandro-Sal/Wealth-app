/**
 * Calculates the Financial Runway (Survival Months).
 * Determines how long the user can survive on current liquid assets based on average monthly expenses.
 * * @param {number|string} year - The year to analyze for expense averages.
 * @return {Object} Object containing runway months and average monthly burn.
 */
function getRunwayData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Retrieve Liquid Net Worth
  // Assumes cell A26 in "Net Worth OGGI" contains the total liquid EUR amount
  const nwSheet = ss.getSheetByName("Net Worth OGGI");
  let liquidCash = 0;
  if (nwSheet) {
    let val = nwSheet.getRange(26, 1).getValue(); // Row 26, Col 1 (EUR)
    liquidCash = (typeof val === 'number') ? val : 0;
  }

  // 2. Calculate Average Monthly Expenses for the selected year
  const expSheetName = "Expenses Tracker " + year;
  const expSheet = ss.getSheetByName(expSheetName);
  
  // Return zero defaults if sheet is missing or empty
  if (!expSheet || expSheet.getLastRow() < 20) {
    return { months: 0, avg: 0 };
  }

  // Read data: Col A (Date), B (Type), and Amount Columns E-J (indices 4-9)
  // Range: From Row 20 to the end, first 10 columns
  const data = expSheet.getRange(20, 1, expSheet.getLastRow() - 19, 10).getValues();
  
  let totalExpense = 0;
  let activeMonths = new Set(); // Tracks unique months with actual spending

  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][0];
    let type = data[i][1];

    // Process only "Expense" rows with valid dates
    if (type === 'Expense' && rowDate instanceof Date) {
      let rowSum = 0;
      // Sum value across bank columns (indices 4 to 9)
      for (let c = 4; c <= 9; c++) {
        let v = parseFloat(data[i][c]);
        if (!isNaN(v)) rowSum += Math.abs(v);
      }

      if (rowSum > 0) {
        totalExpense += rowSum;
        activeMonths.add(rowDate.getMonth()); // Adds month index (0-11) to set
      }
    }
  }

  // Calculate Average
  let monthsCount = activeMonths.size || 1; // Prevent division by zero
  
  let avgMonthly = totalExpense / monthsCount;
  
  // Calculate Runway: Liquid Cash / Average Monthly Expense
  let runway = avgMonthly > 0 ? (liquidCash / avgMonthly) : 0;

  return {
    months: runway.toFixed(1),      // e.g., "12.5"
    avg: avgMonthly.toFixed(0)      // e.g., "1500"
  };
}

/**
 * "Crystal Ball" Projection.
 * Estimates end-of-year savings based on current income/expense trends.
 * For past years, it returns the actual final total.
 * * @param {number|string} year - The year to analyze.
 * @return {Object} Projected total, current total, and status message.
 */
function getYearlyProjection(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expenses Tracker " + year);
  
  // Default values if sheet is missing
  if (!sheet || sheet.getLastRow() < 20) {
    return { projected: "0", current: "0", message: "Insufficient Data" };
  }

  const data = sheet.getRange(20, 1, sheet.getLastRow() - 19, 10).getValues();
  const now = new Date();
  
  // Check if requested year is the current ongoing year
  const isCurrentYear = (String(now.getFullYear()) === String(year));
  
  let income = 0;
  let expense = 0;
  let activeMonths = new Set(); // Track real months with activity

  data.forEach(row => {
    if (!row[0] || !row[1]) return;
    
    // Year Filter
    let d = new Date(row[0]);
    if (d.getFullYear() != year) return;
    
    // Sum Amounts (Cols E-J)
    let amount = 0;
    for (let j = 4; j <= 9; j++) {
      let v = parseFloat(row[j]); 
      if (!isNaN(v)) amount += v; 
    }

    // Classification (Handles Italian keywords for legacy support)
    let type = String(row[1]).toLowerCase();

    if (type.includes('income') || type.includes('entrata')) {
        income += amount;
    } 
    else if (type.includes('expense') || type.includes('spesa') || type.includes('uscita')) {
        expense += Math.abs(amount); // Force positive
    }

    // Mark month as active if there was movement
    if (amount !== 0) activeMonths.add(d.getMonth());
  });

  let currentSavings = income - expense;
  
  // Calculate elapsed months
  let monthsPassed = activeMonths.size || 1; 

  // --- PROJECTION LOGIC ---
  let projected = currentSavings;

  if (isCurrentYear) {
    // Average monthly savings so far
    let avgMonthly = currentSavings / monthsPassed;
    // Remaining months in the year
    let remainingMonths = 12 - monthsPassed;
    
    // Add estimated future savings
    if (remainingMonths > 0) {
      projected = currentSavings + (avgMonthly * remainingMonths);
    }
  }

  return {
    projected: projected.toFixed(0),
    current: currentSavings.toFixed(0),
    message: isCurrentYear ? "Est. Dec 31st" : "Final Actual Total"
  };
}

/**
 * FIRE (Financial Independence, Retire Early) Progress Tracker.
 * Calculates progress % based on Historical/Liquid Net Worth vs. Annual Expenses.
 * Formula: Progress = Capital / (AnnualExpenses * 25)
 * * @param {number|string} year - The reference year.
 * @return {Object} FIRE metrics including Current NW, Target, %, and Bar width.
 */
function getFireProgress(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fireCapital = 0;

  // --- 1. RETRIEVE CAPITAL (Historical from "NW analitico") ---
  const nwSheet = ss.getSheetByName("NW analitico");
  
  if (nwSheet) {
    const headers = nwSheet.getRange(1, 1, 1, nwSheet.getLastColumn()).getValues()[0];
    
    // Find column index for the requested year
    let colIndex = headers.indexOf(String(year));
    if (colIndex === -1) colIndex = headers.findIndex(h => String(h).includes(String(year)));

    if (colIndex > -1) {
       // Read Row 2 (Index 2)
       // use colIndex + 1 because getRange is 1-based
       let val = nwSheet.getRange(2, colIndex + 1).getValue(); 
       
       // Clean data if string formatted
       if (typeof val === 'string') {
         val = parseFloat(val.replace(/[^0-9.,-]/g, '').replace(/\./g, '').replace(',', '.'));
       }
       fireCapital = (typeof val === 'number') ? val : 0;
    } 
  }

  // --- FALLBACK: LIVE DATA (Only if Current Year and historical data missing) ---
  const currentYear = new Date().getFullYear();
  if ((fireCapital === 0 || !fireCapital) && String(year) === String(currentYear)) {
     const sheetOggi = ss.getSheetByName("Net Worth OGGI");
     if (sheetOggi) {
        // Row 26 = Liquid Net Worth (EUR) LIVE
        let val = sheetOggi.getRange(26, 1).getValue(); 
        fireCapital = (typeof val === 'number') ? val : 0;
     }
  }

  // --- 2. CALCULATE ANNUAL EXPENSES ---
  const expSheet = ss.getSheetByName("Expenses Tracker " + year);
  let annualExpenses = 0;
  let currentExpense = 0; // Scope variable correctly
  
  if (expSheet && expSheet.getLastRow() >= 20) {
    const data = expSheet.getRange(20, 1, expSheet.getLastRow() - 19, 10).getValues();
    let activeMonths = new Set();

    data.forEach(row => {
      if (!row[0]) return;
      let d = new Date(row[0]);
      if (d.getFullYear() != year) return; 
      
      let type = String(row[1]).toLowerCase();
      
      if (type.includes('expense') || type.includes('spesa')) {
         let rowSum = 0;
         for (let j = 4; j <= 9; j++) {
           let v = parseFloat(row[j]);
           if (!isNaN(v)) rowSum += Math.abs(v);
         }
         if (rowSum > 0) {
           currentExpense += rowSum;
           activeMonths.add(d.getMonth());
         }
      }
    });

    let months = activeMonths.size || 1;
    
    // Projection Logic:
    // If Current Year -> Project expenses to 12 months based on run rate.
    // If Past Year -> Use actual total expenses.
    if (String(year) === String(currentYear)) {
        annualExpenses = (currentExpense / months) * 12;
    } else {
        annualExpenses = currentExpense;
    }
  }

  if (annualExpenses === 0) annualExpenses = 1; // Prevent division by zero

  // --- 3. CALCULATE FIRE NUMBER ---
  let fireNumber = annualExpenses * 25;
  let progressPct = (fireCapital / fireNumber) * 100;

  // Cap bar at 100% for UI purposes
  let barPct = progressPct > 100 ? 100 : progressPct;

  return {
    currentNW: fireCapital,
    fireTarget: fireNumber,
    pct: progressPct.toFixed(2),
    bar: barPct.toFixed(1)
  };
}