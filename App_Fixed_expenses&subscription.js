/**
 * --- SUBSCRIPTION CONFIGURATION ---
 * Defines the list of recurring expenses (Subscriptions).
 * 'bankCol' refers to the specific column index where the amount should be written.
 * New Optional Properties:
 * - payDay {number}: Day of the month the payment occurs (e.g., 15)
 * - startDate {string}: Format 'YYYY-MM-DD'
 * - endDate {string}: Format 'YYYY-MM-DD' (if present, treats as installment)
 */
/*function _getSubsData() {
  return [
    { id: 0, cat: 'Alloggio', note: 'Affitto', isSplit: true, splits: [{ col: 5, amt: 800 }], payDay: 5 }, 
    { id: 1, cat: 'Free-Time', note: 'Prime Video', amt: 4.99, bankCol: 9, payDay: 15 }, 
    // Installment Example (Corso Inglese):
    { id: 6, cat: 'Altro', note: 'Corso Inglese', amt: 300, bankCol: 5, payDay: 27, startDate: '2025-10-01', endDate: '2026-10-01' }
  ];
}*/

function getSubsList() { return _getSubsData(); }

// ... [mantieni qui la tua funzione addSelectedSubs intatta] ...

/**
 * Calculates the total monthly fixed expenses sum.
 * Filters out subscriptions that are expired or not yet active.
 * @return {number} Total monthly amount.
 */
function getMonthlyFixedCost() {
  const subs = _getSubsData();
  let total = 0;
  const today = new Date();
  
  subs.forEach(sub => {
    let isActive = true;
    if (sub.startDate && today < new Date(sub.startDate)) isActive = false;
    if (sub.endDate && today > new Date(sub.endDate)) isActive = false;

    if (isActive) {
        if (sub.isSplit && sub.splits) {
          sub.splits.forEach(s => total += s.amt);
        } else if (sub.amt) {
          total += sub.amt;
        }
    }
  });
  
  return total;
}

/**
 * Generates a status report for all subscriptions.
 * Calculates remaining payments, sets alerts, and sorts by closest payment date.
 * NOW DYNAMIC: Checks the 'Expenses Tracker' to see if the sub was already paid this month.
 * @return {Array<Object>} Sorted list of subscription statuses.
 */
function getSubsStatus() {
  const subs = _getSubsData();
  const today = new Date();
  const statuses = [];
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();

  // --- 1. Find fixed expenses already paid this month ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Dynamically build the sheet name (e.g., "Expenses Tracker 2026")
  const sheet = ss.getSheetByName("Expenses Tracker " + currentYear);
  let paidNotes = [];

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 20) { // Data starts from row 20
      // Read Date (Col A), Type (Col B), Category (Col C), Note (Col D)
      const data = sheet.getRange(20, 1, lastRow - 19, 4).getValues();
      data.forEach(row => {
        let dateVal = row[0];
        let typeVal = row[1];
        let noteVal = row[3];
        
        // If the row has a valid date belonging to the current month
        if (dateVal instanceof Date && dateVal.getMonth() === currentMonth && dateVal.getFullYear() === currentYear) {
          if (typeVal === "Expense") {
            // Save the note converted to lowercase to avoid case-sensitivity issues
            paidNotes.push(String(noteVal).trim().toLowerCase());
          }
        }
      });
    }
  }

  // Helper to get total days in the current month
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  subs.forEach(sub => {
    let totalAmt = 0;
    if (sub.isSplit && sub.splits) {
      totalAmt = sub.splits.reduce((acc, s) => acc + s.amt, 0);
    } else {
      totalAmt = sub.amt;
    }

    // --- 2. Check if the expense has been paid ---
    const subNoteLower = String(sub.note).trim().toLowerCase();
    const isPaid = paidNotes.includes(subNoteLower);

    let status = {
      note: sub.note,
      amount: totalAmt,
      payDay: sub.payDay || null,
      startDate: sub.startDate || null, 
      endDate: sub.endDate || null,     
      isInstallment: !!sub.endDate,
      isActive: true,
      remainingAmount: 0,
      remainingMonths: 0,
      progressPct: 0,
      alert: false,
      daysUntilNext: 999,
      isPaid: isPaid // <-- NEW PROPERTY EXPORTED TO FRONTEND
    };

    // Check if subscription has started
    if (sub.startDate) {
        const start = new Date(sub.startDate);
        if (today < start) status.isActive = false;
    }

    // Calculate days until next payment
    if (sub.payDay && status.isActive) {
        if (today.getDate() <= sub.payDay) {
            status.daysUntilNext = sub.payDay - today.getDate();
        } else {
            status.daysUntilNext = (daysInMonth - today.getDate()) + sub.payDay;
        }
    }

    // Handle installments (endDate logic)
    if (sub.endDate && status.isActive) {
      const end = new Date(sub.endDate);
      const start = sub.startDate ? new Date(sub.startDate) : new Date();
      
      if (today > end) {
        status.isActive = false; // Expired installment
      } else {
        const totalMonths = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
        const monthsPassed = (today.getFullYear() - start.getFullYear()) * 12 + (today.getMonth() - start.getMonth()) + (today.getDate() >= sub.payDay ? 1 : 0);
        
        status.remainingMonths = Math.max(0, totalMonths - monthsPassed);
        status.remainingAmount = status.remainingMonths * totalAmt;
        status.progressPct = Math.min(100, (monthsPassed / totalMonths) * 100);
        
        if (status.remainingMonths <= 2 && status.remainingMonths > 0) status.alert = "Ending soon!";
      }
    } 
    
    // --- 3. Dynamic Alert Logic (Modified for Paid status) ---
    if (isPaid) {
        status.alert = "Pagato ✅";
        status.daysUntilNext = 9999; // Push to the bottom of the list for better ordering
    } else if (status.isActive && !status.alert && sub.payDay) {
        if (status.daysUntilNext === 0) {
            status.alert = "Due today ⚠️";
        } else if (status.daysUntilNext <= 3) {
            status.alert = "Due in " + status.daysUntilNext + "d ⏳";
        }
    }

    if (status.isActive) statuses.push(status);
  });

  // Sort by days until next payment (ascending order), "Paid" ones will drop to the bottom (9999 days)
  statuses.sort((a, b) => a.daysUntilNext - b.daysUntilNext);

  return statuses;
}

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

/**
 * Calculates the total monthly fixed expenses sum.
 * useful for "Survival Mode" runway calculation.
 * @return {number} Total monthly amount.
 */
function getMonthlyFixedCost() {
  const subs = _getSubsData();
  let total = 0;
  
  subs.forEach(sub => {
    if (sub.isSplit && sub.splits) {
      sub.splits.forEach(s => total += s.amt);
    } else if (sub.amt) {
      total += sub.amt;
    }
  });
  
  return total;
}