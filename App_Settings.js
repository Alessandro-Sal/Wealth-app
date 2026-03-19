/**
 * Fetches application configurations (Categories and Fixed Expenses) from Google Sheets.
 * @returns {Object} { categories: {...}, fixedExpenses: [...] }
 */
function getAppConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 1. FETCH CATEGORIES ---
  const catSheet = ss.getSheetByName("Config_Category");
  let categoriesData = {
    Expense: [],
    Income: [],
    Transfer: [],
    Investment: []
  };

  if (catSheet) {
    const catRows = catSheet.getDataRange().getValues();
    // Start from 1 to skip the header row
    for (let i = 1; i < catRows.length; i++) {
      let type = catRows[i][0];
      let category = catRows[i][1];
      if (type && category) {
        if (!categoriesData[type]) categoriesData[type] = [];
        categoriesData[type].push(category);
      }
    }
  }

  // --- 2. FETCH FIXED EXPENSES ---
  const expSheet = ss.getSheetByName("Config_FixedExpenses");
  let fixedExpensesData = [];

  if (expSheet) {
    const expRows = expSheet.getDataRange().getValues();
    // Start from 1 to skip the header row
    for (let i = 1; i < expRows.length; i++) {
      let row = expRows[i];
      if (row[0] === "") continue; // Skip empty rows

      // Parse amounts carefully (handling '€' and commas ',')
      let rawAmt = String(row[3]).replace('€', '').replace(',', '.').trim();
      let amount = parseFloat(rawAmt) || 0;

      // Handle dates appropriately
      let startDateStr = row[6] ? Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      let endDateStr = row[7] ? Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;

      let isSplitVal = row[8] === true || String(row[8]).toUpperCase() === 'TRUE';

      let sub = {
        id: row[0],
        cat: row[1],
        note: row[2],
        amt: amount,
        bankCol: row[4] ? parseInt(row[4]) : null,
        payDay: row[5] ? parseInt(row[5]) : null,
        startDate: startDateStr,
        endDate: endDateStr,
        isSplit: isSplitVal
      };

      // Handle split logic
      if (sub.isSplit && row[9]) {
        let splits = [];
        let splitParts = String(row[9]).split('|'); // e.g. "5:270,00|7:180,00"
        splitParts.forEach(part => {
          let [col, amtStr] = part.split(':');
          if (col && amtStr) {
            splits.push({
              col: parseInt(col),
              amt: parseFloat(String(amtStr).replace(',', '.'))
            });
          }
        });
        sub.splits = splits;
      }

      fixedExpensesData.push(sub);
    }
  }

  return {
    categories: categoriesData,
    fixedExpenses: fixedExpensesData
  };
}