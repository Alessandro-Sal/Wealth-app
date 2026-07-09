/**
 * Analyzes recent transactions and returns recurring expense suggestions.
 * Uses CacheService to avoid recalculating on every load.
 */
function getSmartSuggestions() {
  // Use a 24-hour cache (86400 seconds) so it's very fast for the user.
  // We don't need real-time updates for suggestions.
  return getFromCache('SMART_SUGGESTIONS', _computeSmartSuggestions, 86400);
}

function _computeSmartSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentYear = new Date().getFullYear();
  const sheet = ss.getSheetByName("Expenses Tracker " + currentYear);
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 20) return [];
  
  // Data starts at row 20
  const data = sheet.getRange(20, 1, lastRow - 19, sheet.getLastColumn()).getValues();
  
  // Filter for the current year (from January 1st)
  const today = new Date();
  const startOfYear = new Date(today.getFullYear(), 0, 1);
  
  // Fetch fixed expenses to exclude them
  let fixedNotes = [];
  try {
    const subs = getAppConfig().fixedExpenses;
    if (subs) {
        fixedNotes = subs.map(s => (s.note || "").toLowerCase().trim());
    }
  } catch(e) {
    console.warn("Could not fetch fixed expenses for exclusion.");
  }
  
  const grouped = {};
  
  data.forEach(row => {
    const date = row[0];
    const type = row[1];
    const category = row[2];
    const note = row[3];
    
    // Ignore non-expenses, empty rows, and older transactions
    if (type !== 'Expense' || !date || !category || !note) return;
    if (new Date(date) < startOfYear) return;
    
    const noteLower = note.toString().toLowerCase().trim();
    if (fixedNotes.includes(noteLower)) return;
    
    // Find the amount and bank column
    // Array index 4 is column E (Cash). Array index 5 is column F.
    let amount = 0;
    let bankColIdx = -1;
    for (let i = 4; i < row.length; i++) {
      if (row[i] && !isNaN(row[i]) && row[i] !== 0) {
        amount = Math.abs(parseFloat(row[i]));
        bankColIdx = i;
        break; // Assume only one bank column used per transaction for simplicity
      }
    }
    
    if (amount === 0) return;
    
    const key = category + "|" + noteLower + "|" + amount;
    
    if (!grouped[key]) {
      grouped[key] = {
        category: category,
        note: note.toString().trim(),
        amount: amount,
        count: 0,
        bankCols: {}
      };
    }
    
    grouped[key].count++;
    if (bankColIdx !== -1) {
      grouped[key].bankCols[bankColIdx] = (grouped[key].bankCols[bankColIdx] || 0) + 1;
    }
  });
  
  // Process groups to find recurring items
  const suggestions = [];
  
  for (const key in grouped) {
    const group = grouped[key];
    
    // Only suggest if it occurred at least 3 times in 90 days
    if (group.count >= 3) {
      // Find most common bank col
      let mostCommonBankCol = -1;
      let maxBankCount = 0;
      for (const bCol in group.bankCols) {
        if (group.bankCols[bCol] > maxBankCount) {
          maxBankCount = group.bankCols[bCol];
          mostCommonBankCol = parseInt(bCol);
        }
      }
      
      suggestions.push({
        category: group.category,
        note: group.note,
        amount: group.amount,
        bankCol: mostCommonBankCol !== -1 ? mostCommonBankCol + 1 : -1, // Convert array index to 1-based column index
        count: group.count
      });
    }
  }
  
  // Sort by frequency descending and return top 5
  suggestions.sort((a, b) => b.count - a.count);
  return suggestions.slice(0, 5);
}
