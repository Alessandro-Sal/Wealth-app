/**
 * Calculates the Financial Runway (Survival Months).
 * Determines how long the user can survive on current liquid assets based on average monthly expenses.
 * Also calculates "Survival Mode" based strictly on fixed recurring costs.
 * * @param {number|string} year - The year to analyze for expense averages.
 * @return {Object} Object containing runway months, average burn, fixed costs, and survival months.
 */
function getRunwayData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Retrieve Liquid Net Worth robustly via Dashboard module (includes Crypto Cache & N/A protection)
  const dashboardData = getDashboardData();
  let liquidCash = dashboardData.rawLiquidNetWorth || 0;

  // 2. Calculate Average Monthly Expenses for the selected year
  const expSheetName = "Expenses Tracker " + year;
  const expSheet = ss.getSheetByName(expSheetName);
  
  // Return zero defaults if sheet is missing or empty
  if (!expSheet || expSheet.getLastRow() < 20) {
    return { months: 0, avg: 0, fixed: 0, survival: 0 };
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
  
  // Calculate Standard Runway: Liquid Cash / Average Monthly Expense
  let runway = avgMonthly > 0 ? (liquidCash / avgMonthly) : 0;

  // --- NEW: Survival Mode Calculation ---
  // Retrieves total fixed costs (subscriptions + recurring) from the helper function
  const fixedCosts = getMonthlyFixedCost(); 
  
  // Calculate Survival Runway: Liquid Cash / Fixed Costs
  // This represents how long you can survive if you cut all discretionary spending
  let survivalRunway = fixedCosts > 0 ? (liquidCash / fixedCosts) : 0;

  return {
    months: runway.toFixed(1),      // e.g., "12.5" (Standard based on lifestyle)
    avg: avgMonthly.toFixed(0),     // e.g., "1500" (Real Monthly Average)
    fixed: fixedCosts.toFixed(0),       // Total Fixed Expenses
    survival: survivalRunway.toFixed(1) // Survival Months (Emergency Mode)
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
 * @param {number|string} year - The reference year.
 * @return {Object} FIRE metrics including Current NW, Target, %, and Bar width.
 */
function getFireProgress(year) {
  console.log("--- STARTING getFireProgress FOR YEAR: " + year + " ---");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fireCapital = 0;
  const currentYear = new Date().getFullYear();

  // --- 1. RETRIEVE CAPITAL (Strictly Historical from "NW Analitico") ---
  const nwSheet = ss.getSheetByName("NW Analitico");
  
  if (nwSheet) {
    const headers = nwSheet.getRange(1, 1, 1, nwSheet.getLastColumn()).getValues()[0];
    
    // Strict match to find the exact year column (ignoring monthly date objects)
    let colIndex = headers.findIndex(h => {
      if (!h) return false;
      let headerStr = String(h).trim();
      return headerStr === String(year);
    });

    console.log("DEBUG: Exact Column Index for year " + year + " is:", colIndex);

    if (colIndex > -1) {
       let val = nwSheet.getRange(2, colIndex + 1).getValue(); 
       console.log("DEBUG: Raw NW value from sheet (Row 2, Col " + (colIndex + 1) + "):", val);
       
       // Clean data if string formatted
       if (typeof val === 'string') {
         val = parseFloat(val.replace(/[^0-9.,-]/g, '').replace(/\./g, '').replace(',', '.'));
       }
       fireCapital = (typeof val === 'number' && !isNaN(val)) ? val : 0;
    } 
  } else {
    console.log("DEBUG: Sheet 'NW analitico' NOT FOUND!");
  }

  console.log("DEBUG: FINAL fireCapital to be used:", fireCapital);

  // --- 2. CALCULATE ANNUAL EXPENSES ---
  // Try to use past year for a stable FIRE target
  let expenseYear = String(year) === String(currentYear) ? parseInt(year) - 1 : parseInt(year);
  let expSheet = ss.getSheetByName("Expenses Tracker " + expenseYear);
  let isFallbackYear = false;

  console.log("DEBUG: Trying to fetch expenses for year:", expenseYear);

  // Fallback: If the past year's sheet doesn't exist, use requested year
  if (!expSheet) {
      console.log("DEBUG: Sheet 'Expenses Tracker " + expenseYear + "' NOT FOUND. Falling back to current year.");
      expenseYear = parseInt(year);
      expSheet = ss.getSheetByName("Expenses Tracker " + expenseYear);
      isFallbackYear = true;
  }

  let annualExpenses = 0;
  let currentExpense = 0;
  let activeMonths = new Set();
  
  if (expSheet && expSheet.getLastRow() >= 20) {
    const data = expSheet.getRange(20, 1, expSheet.getLastRow() - 19, 10).getValues();

    data.forEach(row => {
      if (!row[0]) return;
      
      let d = new Date(row[0]);
      // Validate date object and check if it matches the target year
      if (isNaN(d.getTime()) || d.getFullYear() != expenseYear) return; 
      
      let type = String(row[1]).toLowerCase();
      
      if (type.includes('expense') || type.includes('spesa')) {
         let rowAmount = 0;
         
         // Sum columns E to J (Indices 4 to 9) allowing positive/negative netting
         for (let j = 4; j <= 9; j++) {
           let val = row[j];
           
           if (typeof val === 'string') {
               let cleanStr = val.replace(/[^0-9.,-]/g, '');
               if (cleanStr.includes(',') && cleanStr.includes('.')) {
                   cleanStr = cleanStr.replace(/\./g, '').replace(',', '.'); 
               } else if (cleanStr.includes(',')) {
                   cleanStr = cleanStr.replace(',', '.');
               }
               val = parseFloat(cleanStr);
           }
           if (typeof val === 'number' && !isNaN(val)) {
               rowAmount += val; 
           }
         }
         
         // Apply absolute value AFTER the row is summed (Matches getYearlyTotals logic)
         if (rowAmount !== 0) {
            currentExpense += Math.abs(rowAmount);
            activeMonths.add(d.getMonth());
         }
      }
    });

    console.log("DEBUG: Total raw expenses found:", currentExpense, "| Active months:", activeMonths.size);

    if (isFallbackYear && String(year) === String(currentYear)) {
        let months = activeMonths.size || 1;
        annualExpenses = (currentExpense / months) * 12;
        console.log("DEBUG: Projected annual expenses (Run Rate):", annualExpenses);
    } else {
        annualExpenses = currentExpense;
        console.log("DEBUG: Historical annual expenses:", annualExpenses);
    }
  } else {
    console.log("DEBUG: Expense sheet missing or empty!");
  }

  if (annualExpenses === 0) annualExpenses = 1; // Prevent division by zero

  // --- 3. CALCULATE FIRE NUMBER ---
  let fireNumber = annualExpenses * 25;
  let progressPct = (fireCapital / fireNumber) * 100;

  // Cap bar at 100% for UI purposes
  let barPct = progressPct > 100 ? 100 : progressPct;

  console.log("DEBUG: Final FIRE Target:", fireNumber);
  console.log("DEBUG: Final Progress %:", progressPct);
  console.log("--- END getFireProgress ---");

  return {
    currentNW: fireCapital,
    fireTarget: fireNumber,
    pct: progressPct.toFixed(2),
    bar: barPct.toFixed(1)
  };
}