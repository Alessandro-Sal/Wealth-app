/**
 * --- MONTHLY SAVINGS FLOW ---
 * Calculates financial metrics for the CURRENT month.
 * Scans the expense tracker to sum Income and Expenses, then computes
 * the Net Savings and Savings Rate %.
 * * @return {Object} Object containing formatted strings for income, expenses, savings, and rate.
 */
function getMonthlySavings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Target specific year sheet
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  
  const lastRow = sheet.getLastRow();
  
  // If no data exists (assuming data starts at Row 20), return zero values
  if (lastRow < 20) {
       return { income: "0.00", expenses: "0.00", savings: "0.00", rate: "0.0" };
  }
  
  // Read data from Row 20 to the end, first 10 columns
  const data = sheet.getRange(20, 1, lastRow-19, 10).getValues();
  
  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();
  
  let income = 0; 
  let expenses = 0;
  
  data.forEach(row => {
    // Ensure row has valid Date (Col A) and Type (Col B)
    if(row[0] && row[1]) {
      let d = new Date(row[0]);
      
      // Filter: Current Month & Year only
      if(d.getMonth() === curMonth && d.getFullYear() === curYear) {
        let amt = 0;
        // Sum values across bank columns (Indices 4-9 / Cols E-J)
        for(let i=4; i<=9; i++) { 
           let val = parseFloat(row[i]);
           if(!isNaN(val)) amt += val; 
        }
        
        // Aggregate totals based on transaction type
        if(row[1] === 'Income') income += amt;
        if(row[1] === 'Expense') expenses += Math.abs(amt); 
      }
    }
  });
  
  let savings = income - expenses;
  // Calculate Savings Rate (prevent division by zero)
  let rate = income > 0 ? (savings / income) * 100 : 0;
  
  return { 
    income: income.toFixed(2), 
    expenses: expenses.toFixed(2), 
    savings: savings.toFixed(2), 
    rate: rate.toFixed(1) 
  };
}