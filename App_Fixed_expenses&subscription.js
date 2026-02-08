/**
 * --- SUBSCRIPTION CONFIGURATION ---
 * Defines the list of recurring expenses (Subscriptions).
 * 'bankCol' refers to the specific column index where the amount should be written.
 * Note: Categories ('cat') are kept in Italian to match Spreadsheet validation rules.
 */
function _getSubsData() {
  return [
    { 
      id: 0, 
      cat: 'Alloggio', 
      note: 'Affitto (Intesa + Wallet)', 
      isSplit: true, 
      splits: [{ col: 5, amt: 270.00 }, { col: 7, amt: 180.00 }]
    }, 
    { id: 1, cat: 'Free-Time', note: 'Prime Video', amt: 4.99, bankCol: 9 }, 
    { id: 2, cat: 'Necessità', note: 'iCloud', amt: 2.99, bankCol: 9 }, 
    { id: 3, cat: 'Necessità', note: 'Phone Top-up', amt: 9.99, bankCol: 6 },  
    { id: 4, cat: 'Free-Time', note: 'Spotify', amt: 6.49, bankCol: 6 },
    { id: 5, cat: 'Altro', note: 'Corso Data Analyst', amt: 600.00, bankCol: 5 },
    { id: 6, cat: 'Altro', note: 'Corso Inglese', amt: 165.00, bankCol: 5 }
  ];
}

/**
 * Public getter for the subscription list.
 * @return {Array<Object>} The array of subscription objects.
 */
function getSubsList() { return _getSubsData(); }

/**
 * Adds selected subscriptions to the "Expenses Tracker" sheet.
 * Locates the first available empty row (starting from row 20) to avoid overwriting data.
 * Handles both single-payment and split-payment subscriptions automatically.
 * * @param {Array<number|string>} selectedIds - Array of IDs corresponding to the subscriptions to add.
 * @return {string} Status message indicating success or error.
 */
function addSelectedSubs(selectedIds) {
  const allSubs = _getSubsData();
  // Ensure IDs are numbers for comparison
  const idsNumbers = selectedIds.map(id => Number(id));
  const subsToAdd = allSubs.filter(sub => idsNumbers.includes(sub.id));

  if (subsToAdd.length === 0) return "No expenses selected.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Target Sheet: Expenses Tracker 2026
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  if (!sheet) return "Error: Sheet not found";

  const startRow = 20;
  const lastRow = sheet.getLastRow();
  let newRow = startRow;

  // --- ROW FINDING LOGIC ---
  // Scans Column B starting from row 20 to find the first truly empty row.
  if (lastRow >= startRow) {
    const rowsCheck = lastRow - startRow + 1;
    if (rowsCheck > 0) {
      const colB = sheet.getRange(startRow, 2, rowsCheck, 1).getValues();
      for (let i = 0; i < colB.length; i++) {
        // If cell is empty, this is our target row
        if (!colB[i][0] || colB[i][0] === "") { 
            newRow = startRow + i; 
            break; 
        }
        // If we reached the end of the data, append after the last one
        if (i === colB.length - 1) newRow = startRow + i + 1;
      }
    }
  }

  const today = new Date();
  
  // --- WRITE DATA ---
  subsToAdd.forEach(sub => {
    // If the calculated row is beyond the sheet max rows, add a new row
    if (newRow > sheet.getMaxRows()) sheet.insertRowAfter(sheet.getMaxRows());
    
    // Set Metadata
    sheet.getRange(newRow, 1).setValue(today);
    sheet.getRange(newRow, 2).setValue("Expense");
    sheet.getRange(newRow, 3).setValue(sub.cat);
    sheet.getRange(newRow, 4).setValue(sub.note);

    // Handle Split Payments vs Single Column Payment
    if (sub.isSplit && sub.splits) {
        sub.splits.forEach(splitItem => {
            // Write negative amount to the specific column defined in the split object
            sheet.getRange(newRow, splitItem.col).setValue(-Math.abs(splitItem.amt));
        });
    } else {
        // Write negative amount to the standard bank column
        sheet.getRange(newRow, sub.bankCol).setValue(-Math.abs(sub.amt));
    }
    
    // Move to next row for the next subscription in the loop
    newRow++;
  });

  return "Added: " + subsToAdd.map(s => s.note).join(", ");
}