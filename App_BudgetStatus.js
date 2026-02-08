/**
 * Calculates the current month's budget status against defined limits.
 * Iterates through the specific "Expenses Tracker 2026" sheet to sum expenses by category.
 * Filters for the current month/year and returns a sorted list of categories based on percentage used.
 * * @return {Array<Object>} Array of budget objects {category, spent, limit, pct}, sorted by highest usage %.
 */
function getBudgetStatus() {
  // Define hardcoded monthly budget limits (Categories kept in Italian to match Sheet data)
  const BUDGET_LIMITS = {
    'Alimentazione': 150, 'Necessità': 150, 'Free-Time': 300, 'Alloggio': 450,
    'Regali': 50, 'Uscite': 150, 'Trasporti': 100, 'Viaggi': 300, 'Varie': 50
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Access the specific tracker sheet
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  
  if (!sheet) return []; // Return empty if sheet missing

  const startRow = 20; 
  const lastRow = sheet.getLastRow();

  // If no data exists beyond the header/summary section, return initial zeroed structure
  if (lastRow < startRow) {
    return Object.keys(BUDGET_LIMITS).map(c => ({ category: c, spent: 0, limit: BUDGET_LIMITS[c], pct: 0 }));
  }
  
  // Fetch transaction data (Columns A to J)
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues();
  
  const now = new Date(); 
  const currentMonth = now.getMonth(); 
  const currentYear = now.getFullYear();
  
  // Initialize spending accumulators
  let spending = {}; 
  for (let cat in BUDGET_LIMITS) { spending[cat] = 0; }
  
  // Process transactions
  for (let i = 0; i < data.length; i++) {
    let rowDate = new Date(data[i][0]); 
    let type = data[i][1]; 
    let cat = String(data[i][2]).trim();

    // Filter: Must be an 'Expense' and fall within the current month/year
    if (type === 'Expense' && rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      let amount = 0; 
      // Sum columns 4 through 9 (indices 4-9) representing different payment accounts/methods
      for (let j = 4; j <= 9; j++) { 
        let val = parseFloat(data[i][j]); 
        if (!isNaN(val)) amount += val; 
      }
      
      if (spending.hasOwnProperty(cat)) {
         spending[cat] += Math.abs(amount);
      }
    }
  }

  // Calculate percentages and format result
  let result = []; 
  for (let cat in BUDGET_LIMITS) { 
    let spent = spending[cat]; 
    let limit = BUDGET_LIMITS[cat]; 
    result.push({ 
      category: cat, 
      spent: spent, 
      limit: limit, 
      pct: (spent / limit) * 100 
    }); 
  }

  // Sort by highest percentage used (Critical budgets first)
  return result.sort((a, b) => b.pct - a.pct);
}