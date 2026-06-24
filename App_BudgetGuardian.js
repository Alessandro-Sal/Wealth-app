/**
 * App_BudgetGuardian.js
 * Monitors CURRENT MONTH's expenses.
 * Trigger: Runs weekly (e.g., Friday).
 * Logic: Alerts if Burn Rate > 70% or if Net Savings is negative for the current month.
 */

function checkBudgetStatus() {
  // 1. Get detailed analysis for the CURRENT MONTH
  const monthlyData = getCurrentMonthAnalysis();

  // 2. Calculate Total Budget for Reference
  // L'utente ha richiesto esplicitamente di usare 1800 come budget di riferimento
  const totalBudget = 1800;

  // Burn rate is Expense against Total Budget
  const burnRate = (monthlyData.expense / totalBudget) * 100;

  // TRIGGER CONDITION:
  // A. Burn Rate > 70% (Spending too fast against the budget)
  if (burnRate > 70) {
    
    console.log(`Budget Alert Triggered. Burn Rate: ${burnRate.toFixed(1)}%`);
    
    // Generate AI Scolding
    const aiWarning = generateBudgetAI(monthlyData, burnRate);
    
    MailApp.sendEmail({
      to: getNotifyEmail(), // FIX (2.17): email centralizzata (vedi App_NotifyConfig.js)
      subject: `⚠️ ALERT BUDGET: Burn Rate al ${burnRate.toFixed(1)}%`,
      htmlBody: `<div style="font-family: Arial, sans-serif; padding: 20px; background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 8px; color: #856404;">
        <h2 style="margin-top: 0;">👮‍♂️ Budget Guardian Alert</h2>
        <p><strong>Mese Corrente:</strong> Entrate €${monthlyData.income.toFixed(0)} | Uscite €${monthlyData.expense.toFixed(0)}</p>
        <hr style="border-top: 1px solid #ffeeba;">
        ${aiWarning}
      </div>`
    });
  } else {
    console.log(`Budget Safe. Burn Rate: ${burnRate.toFixed(1)}%`);
  }
}

/**
 * Helper: Analyzes the specific sheet for the current month.
 * Returns totals AND top spending categories.
 */
function getCurrentMonthAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-11
  
  const sheetName = `Expenses Tracker ${currentYear}`;
  const sheet = ss.getSheetByName(sheetName);
  
  // Default structure
  let result = { income: 0, expense: 0, topCategories: [] };
  
  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow < 20) return result; // Assuming data starts at row 20

  // Read Columns A (Date) to J (Banks) - adjust range as per your sheet
  const data = sheet.getRange(20, 1, lastRow - 19, 10).getValues();
  
  let expenseMap = {};

  data.forEach(row => {
    // Row[0] = Date, Row[1] = Type, Row[2] = Category
    if (row[0] && row[1]) {
      let d = new Date(row[0]);
      
      // FILTER: Only process rows belonging to the CURRENT MONTH & YEAR
      if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
        
        // Sum amount from Bank Columns (Index 4 to 9)
        let amount = 0;
        for (let i = 4; i <= 9; i++) {
          let val = parseFloat(row[i]);
          if (!isNaN(val)) amount += val;
        }
        
        let type = String(row[1]);
        let category = String(row[2]).trim();

        if (type === 'Income') {
          result.income += amount;
        } else if (type === 'Expense') {
          // Expenses are usually negative in DB, we treat them as absolute for burn rate
          let absAmount = Math.abs(amount);
          result.expense += absAmount;
          
          // Map category for AI analysis
          if (category) {
            expenseMap[category] = (expenseMap[category] || 0) + absAmount;
          }
        }
      }
    }
  });

  // Sort Categories by spend (Desc) and take Top 3
  result.topCategories = Object.entries(expenseMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([cat, val]) => `${cat} (€${val.toFixed(0)})`);

  return result;
}

/**
 * Generates the warning message using AI with Fallback.
 */
function generateBudgetAI(data, rate) {
  // 1. Calculate End-of-Month Projection
  const today = new Date();
  const dayOfMonth = today.getDate();
  // Get total days in current month (day 0 of next month is the last day of current)
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  
  // Linear projection: (Current Expense / Days Passed) * Total Days
  const projectedExpense = (data.expense / dayOfMonth) * daysInMonth;

  const prompt = `
    ROLE: You are a strict personal accountant.
    CONTEXT: Current Month Analysis.
    
    TIME CONTEXT: Day ${dayOfMonth} of ${daysInMonth}.
    
    DATA: 
    - Income: €${data.income.toFixed(0)}
    - Expenses (Current): €${data.expense.toFixed(0)}
    - PROJECTED EXPENSE (End of Month): €${projectedExpense.toFixed(0)}
    - Burn Rate: ${rate.toFixed(1)}% (Threshold is 70%).
    - TOP SPENDING CATEGORIES THIS MONTH: ${JSON.stringify(data.topCategories)}
    
    TASK:
    Scold the user (in Italian, friendly but firm).
    1. Highlight the gap between Current Expense and Projected Expense if relevant.
    2. Explicitly mention the category where they are wasting the most money.
    3. Give one specific, actionable tip to save money for the remaining ${daysInMonth - dayOfMonth} days.
    Keep it short (max 3-4 sentences).
  `;

  let response = fetchUniversalAI(prompt, 'OPENROUTER');
  if (!response) {
      response = fetchUniversalAI(prompt, 'GEMINI', true);
  }
  
  return response || "<p>AI unavailable. Stop spending money!</p>";
}