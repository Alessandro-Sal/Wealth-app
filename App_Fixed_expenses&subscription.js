/**
 * Ottiene la lista delle spese fisse dal foglio di configurazione.
 */
function getSubsList() { 
  return getAppConfig().fixedExpenses; 
}

/**
 * Calcola il totale mensile delle spese fisse (per la Dashboard / Runway).
 */
function getMonthlyFixedCost() {
  const subs = getAppConfig().fixedExpenses;
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
 * Genera lo stato di tutte le iscrizioni (Pagato, In Scadenza, ecc.)
 */
function getSubsStatus() {
  const subs = getAppConfig().fixedExpenses;
  const today = new Date();
  const statuses = [];
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expenses Tracker " + currentYear);
  let paidNotes = [];

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 20) {
      const data = sheet.getRange(20, 1, lastRow - 19, 4).getValues();
      data.forEach(row => {
        let dateVal = row[0];
        let typeVal = row[1];
        let noteVal = row[3];
        
        if (dateVal instanceof Date && dateVal.getMonth() === currentMonth && dateVal.getFullYear() === currentYear) {
          if (typeVal === "Expense") paidNotes.push(String(noteVal).trim().toLowerCase());
        }
      });
    }
  }

  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  subs.forEach(sub => {
    let totalAmt = (sub.isSplit && sub.splits) ? sub.splits.reduce((acc, s) => acc + s.amt, 0) : sub.amt;
    const isPaid = paidNotes.includes(String(sub.note).trim().toLowerCase());

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
      isPaid: isPaid 
    };

    if (sub.startDate && today < new Date(sub.startDate)) status.isActive = false;

    if (sub.payDay && status.isActive) {
        if (today.getDate() <= sub.payDay) status.daysUntilNext = sub.payDay - today.getDate();
        else status.daysUntilNext = (daysInMonth - today.getDate()) + sub.payDay;
    }

    if (sub.endDate && status.isActive) {
      const end = new Date(sub.endDate);
      const start = sub.startDate ? new Date(sub.startDate) : new Date();
      if (today > end) {
        status.isActive = false; 
      } else {
        const totalMonths = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
        const monthsPassed = (today.getFullYear() - start.getFullYear()) * 12 + (today.getMonth() - start.getMonth()) + (today.getDate() >= sub.payDay ? 1 : 0);
        
        status.remainingMonths = Math.max(0, totalMonths - monthsPassed);
        status.remainingAmount = status.remainingMonths * totalAmt;
        status.progressPct = Math.min(100, (monthsPassed / totalMonths) * 100);
        if (status.remainingMonths <= 2 && status.remainingMonths > 0) status.alert = "Ending soon!";
      }
    } 
    
    if (isPaid) {
        status.alert = "Pagato ✅";
        status.daysUntilNext = 9999; 
    } else if (status.isActive && !status.alert && sub.payDay) {
        if (status.daysUntilNext === 0) status.alert = "Due today ⚠️";
        else if (status.daysUntilNext <= 3) status.alert = "Due in " + status.daysUntilNext + "d ⏳";
    }

    if (status.isActive) statuses.push(status);
  });

  statuses.sort((a, b) => a.daysUntilNext - b.daysUntilNext);
  return statuses;
}

/**
 * Aggiunge le spese selezionate all'Expenses Tracker (OPTIMIZED)
 */
function addSelectedSubs(selectedIds) {
  const allSubs = getAppConfig().fixedExpenses;
  const idsNumbers = selectedIds.map(id => Number(id));
  const subsToAdd = allSubs.filter(sub => idsNumbers.includes(Number(sub.id)));

  if (subsToAdd.length === 0) return "Nessuna spesa trovata nel database.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = new Date().getFullYear(); 
  const sheet = ss.getSheetByName("Expenses Tracker " + year);
  if (!sheet) return "Errore: Foglio Expenses Tracker " + year + " non trovato";

  const startRow = 20;
  const lastRow = sheet.getLastRow();
  let newRow = startRow;

  if (lastRow >= startRow) {
    const colB = sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();
    for (let i = 0; i < colB.length; i++) {
      if (!colB[i][0] || colB[i][0] === "") { newRow = startRow + i; break; }
      if (i === colB.length - 1) newRow = startRow + i + 1;
    }
  }

  const today = new Date();
  const maxCols = sheet.getLastColumn();
  const rowsToAppend = []; // Array per impacchettare tutte le spese
  
  subsToAdd.forEach(sub => {
    let rowData = new Array(maxCols).fill("");
    
    rowData[0] = today;           // A: Date
    rowData[1] = "Expense";       // B: Type
    rowData[2] = sub.cat;         // C: Category
    rowData[3] = sub.note;        // D: Note

    if (sub.isSplit && sub.splits) {
        sub.splits.forEach(splitItem => {
            if (splitItem.col > 0) rowData[splitItem.col - 1] = -Math.abs(splitItem.amt);
        });
    } else {
        if (sub.bankCol > 0) rowData[sub.bankCol - 1] = -Math.abs(sub.amt);
    }
    
    rowsToAppend.push(rowData);
  });

  // Scrive l'intero blocco di spese in una singola chiamata API
  if (rowsToAppend.length > 0) {
    // Se serve, aggiunge righe vuote prima di incollare
    if (newRow + rowsToAppend.length - 1 > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), rowsToAppend.length);
    }
    sheet.getRange(newRow, 1, rowsToAppend.length, maxCols).setValues(rowsToAppend);
  }

  return "Added: " + subsToAdd.map(s => s.note).join(", ");
}

/**
 * Aggiunge una nuova Spesa Fissa dal menu Settings
 */
function addFixedExpenseToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config_FixedExpenses");
  if (!sheet) throw new Error("Foglio Config_FixedExpenses mancante");
  
  // Genera un nuovo ID progressivo
  const lastRow = sheet.getLastRow();
  let newId = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    newId = Math.max(...ids.map(id => Number(id[0]) || 0)) + 1;
  }
  
  // Scrive la riga includendo i parametri dello split
  sheet.appendRow([
    newId, 
    data.cat, 
    data.note, 
    data.amt, 
    data.bankCol || "", // Se è split, la banca principale è vuota
    data.payDay || "", 
    data.startDate || "", 
    data.endDate || "", 
    data.isSplit || false, // TRUE se l'interruttore è stato attivato
    data.splitDetails || "" // Es. "5:270|7:180"
  ]);
  
  return `Spesa ${data.note} aggiunta!`;
}