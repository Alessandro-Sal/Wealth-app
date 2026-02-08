/**
 * Aggregates monthly totals for Income, Expenses, and Savings.
 * Used to populate the main bar charts.
 * Calculates dynamic averages: uses 12 months for past years, or year-to-date months for the current year.
 * * @param {number|string} year - The year to analyze.
 * @return {Object} Object containing arrays (size 12) for income, expense, savings, and their averages.
 */
function getMonthlyChartData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Expenses Tracker " + year;
  const sheet = ss.getSheetByName(sheetName);
  
  let incomeData = new Array(12).fill(0);
  let expenseData = new Array(12).fill(0);
  let savingsData = new Array(12).fill(0);
  
  // If sheet is missing or empty, return zeroed data
  if (!sheet || sheet.getLastRow() < 20) {
    return { 
      income: incomeData, 
      expense: expenseData, 
      savings: savingsData, 
      averages: { income: 0, expense: 0, savings: 0 } 
    };
  }

  const lastRow = sheet.getLastRow();
  // Read data starting from Row 20
  const data = sheet.getRange(20, 1, lastRow - 19, 10).getValues();
  
  data.forEach(row => {
    if (row[0] && row[1]) {
      let date = new Date(row[0]);
      if (!isNaN(date) && date.getFullYear() == year) {
        let monthIndex = date.getMonth(); 
        let amount = 0;
        
        // Sum Bank Columns (Indices 4-9)
        for (let i = 4; i <= 9; i++) {
          let val = parseFloat(row[i]);
          if (!isNaN(val)) amount += val;
        }

        if (row[1] === 'Income') {
           incomeData[monthIndex] += amount;
        } else if (row[1] === 'Expense') {
           expenseData[monthIndex] += Math.abs(amount);
        }
      }
    }
  });

  let totInc = 0;
  let totExp = 0;
  
  for (let i = 0; i < 12; i++) {
    savingsData[i] = incomeData[i] - expenseData[i];
    totInc += incomeData[i];
    totExp += expenseData[i];
  }
  
  // --- AVERAGE DIVISOR CALCULATION ---
  // Use 12 for past years.
  // Use actual months passed (current month + 1) for the current year.
  let divisor = 12;
  const now = new Date();
  if (parseInt(year) === now.getFullYear()) {
      divisor = now.getMonth() + 1;
      if(divisor < 1) divisor = 1; 
  }

  return { 
    income: incomeData, 
    expense: expenseData, 
    savings: savingsData,
    averages: {
      income: totInc / divisor,
      expense: totExp / divisor,
      savings: (totInc - totExp) / divisor
    }
  };
}

/**
 * Analyzes spending and income patterns broken down by Category.
 * Returns monthly data maps and calculated averages for each specific category.
 * Useful for "Top Expenses" visualization or category drill-down.
 * * @param {number|string} year - The year to analyze.
 * @return {Object} Object containing category maps and detailed average metrics.
 */
function getMonthlyCategoryBreakdown(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Expenses Tracker " + year;
  const sheet = ss.getSheetByName(sheetName);

  let incomeMap = {};
  let expenseMap = {};
  
  let totalIncomeVal = 0;
  let totalExpenseVal = 0;

  if (!sheet || sheet.getLastRow() < 20) {
    return { income: {}, expense: {}, avgIncome: 0, avgExpense: 0, catAverages: { income: {}, expense: {} } };
  }

  const data = sheet.getRange(20, 1, sheet.getLastRow() - 19, 10).getValues();

  data.forEach(row => {
    if (!row[0] || !row[1]) return;
    let date = new Date(row[0]);
    
    if (!isNaN(date) && date.getFullYear() == year) {
      let month = date.getMonth();
      let type = row[1];
      let cat = String(row[2]).trim();
      
      let amount = 0;
      for (let i = 4; i <= 9; i++) {
        let val = parseFloat(row[i]);
        if (!isNaN(val)) amount += val;
      }
      
      if (amount === 0) return;

      if (type === 'Expense') {
        amount = Math.abs(amount);
        if (!expenseMap[cat]) expenseMap[cat] = new Array(12).fill(0);
        expenseMap[cat][month] += amount;
        totalExpenseVal += amount;
      } 
      else if (type === 'Income') {
        if (!incomeMap[cat]) incomeMap[cat] = new Array(12).fill(0);
        incomeMap[cat][month] += amount;
        totalIncomeVal += amount;
      }
    }
  });

  // Calculate Divisor (Months passed)
  let divisor = 12;
  const now = new Date();
  if (parseInt(year) === now.getFullYear()) {
      divisor = now.getMonth() + 1;
      if(divisor < 1) divisor = 1;
  }

  // --- NEW: Calculate averages per single category ---
  let catAverages = { income: {}, expense: {} };

  // Expense Averages
  for (let cat in expenseMap) {
    let sum = expenseMap[cat].reduce((a, b) => a + b, 0);
    catAverages.expense[cat] = sum / divisor;
  }
  // Income Averages
  for (let cat in incomeMap) {
    let sum = incomeMap[cat].reduce((a, b) => a + b, 0);
    catAverages.income[cat] = sum / divisor;
  }

  return { 
    income: incomeMap, 
    expense: expenseMap,
    avgIncome: totalIncomeVal / divisor,
    avgExpense: totalExpenseVal / divisor,
    catAverages: catAverages // Return detailed averages
  };
}