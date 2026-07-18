/**
 * Assicura che il foglio Config_PlannedExpenses esista.
 */
function ensurePlannedExpensesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Config_PlannedExpenses");
  if (!sheet) {
    sheet = ss.insertSheet("Config_PlannedExpenses");
    sheet.appendRow(["ID", "Category", "Note", "Amount", "Expected Month (1-12)", "Expected Year", "Bank Column", "Expected Day"]);
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Legge le spese programmate e ne calcola lo stato.
 */
function getPlannedExpensesStatus() {
  const sheet = ensurePlannedExpensesSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const currentYear = new Date().getFullYear();
  
  // Trova le spese pagate quest'anno
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expSheet = ss.getSheetByName("Expenses Tracker " + currentYear);
  let paidNotes = [];
  
  if (expSheet && expSheet.getLastRow() >= 20) {
    const expData = expSheet.getRange(20, 1, expSheet.getLastRow() - 19, 4).getValues();
    expData.forEach(r => {
      let typeVal = r[1];
      let noteVal = r[3];
      if (typeVal === "Expense" && noteVal) {
        paidNotes.push(String(noteVal).replace(/\s+/g, ' ').trim().toLowerCase());
      }
    });
  }

  let planned = [];
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (!row[0]) continue;
    
    let note = row[2];
    let cleanNote = String(note).replace(/\s+/g, ' ').trim().toLowerCase();
    
    let expectedYear = parseInt(row[5]) || currentYear;
    // E' pagata se l'anno è quello corrente o se per qualche motivo abbiamo salvato il pagamento 
    let isPaid = false;
    if (expectedYear === currentYear) {
      isPaid = paidNotes.includes(cleanNote);
    }

    planned.push({
      id: row[0],
      cat: row[1],
      note: note,
      amount: parseFloat(row[3]) || 0,
      expectedMonth: parseInt(row[4]) || 1, // 1-12
      expectedYear: expectedYear,
      bankCol: row[6] || "",
      expectedDay: parseInt(row[7]) || null,
      isPaid: isPaid
    });
  }

  // Sort by Year, Month
  planned.sort((a, b) => {
    if (a.expectedYear !== b.expectedYear) return a.expectedYear - b.expectedYear;
    return a.expectedMonth - b.expectedMonth;
  });

  return planned;
}

/**
 * Calcola il totale delle spese programmate (NON pagate) per l'anno in corso.
 * Serve per impattare il Runway / Budget.
 */
function getUnpaidPlannedExpensesTotal(year) {
  const planned = getPlannedExpensesStatus();
  let total = 0;
  planned.forEach(p => {
    if (p.expectedYear === year && !p.isPaid) {
      total += p.amount;
    }
  });
  return total;
}

function addPlannedExpense(data) {
  const sheet = ensurePlannedExpensesSheet();
  const lastRow = sheet.getLastRow();
  let newId = 1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    newId = Math.max(...ids.map(id => Number(id[0]) || 0)) + 1;
  }
  
  sheet.appendRow([
    newId,
    data.cat,
    data.note,
    data.amount,
    data.expectedMonth,
    data.expectedYear,
    data.bankCol || "",
    data.expectedDay || ""
  ]);
  
  return "Spesa programmata '" + data.note + "' salvata!";
}

function payPlannedExpense(id, actualAmount, bankCol) {
  const planned = getPlannedExpensesStatus();
  const exp = planned.find(p => String(p.id) === String(id));
  if (!exp) throw new Error("Spesa non trovata.");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentYear = new Date().getFullYear();
  const sheet = ss.getSheetByName("Expenses Tracker " + currentYear);
  if (!sheet) throw new Error("Foglio Expenses Tracker " + currentYear + " non trovato");

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

  const maxCols = sheet.getLastColumn();
  let rowData = new Array(maxCols).fill("");
  
  rowData[0] = new Date();           // A: Date
  rowData[1] = "Expense";            // B: Type
  rowData[2] = exp.cat;              // C: Category
  rowData[3] = exp.note;             // D: Note
  
  let useBank = bankCol ? parseInt(bankCol) : (exp.bankCol ? parseInt(exp.bankCol) : 5); // Fallback to col E (bank 1)
  const amt = actualAmount !== null && actualAmount !== undefined && actualAmount !== "" ? parseFloat(actualAmount) : exp.amount;
  if (useBank > 0) rowData[useBank - 1] = -Math.abs(amt);

  if (newRow > sheet.getMaxRows()) {
    sheet.insertRowAfter(sheet.getMaxRows());
  }
  sheet.getRange(newRow, 1, 1, maxCols).setValues([rowData]);
  
  // Rimuovi la spesa programmata dal foglio dopo averla pagata
  deletePlannedExpense(id);
  
  return "Spesa '" + exp.note + "' registrata come pagata ed eliminata dal modulo!";
}

function deletePlannedExpense(id) {
  const sheet = ensurePlannedExpensesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return "Spesa programmata eliminata.";
    }
  }
  throw new Error("Spesa non trovata.");
}

function editPlannedExpense(id, data) {
  const sheet = ensurePlannedExpensesSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 7).setValues([[
        data.cat,
        data.note,
        data.amount,
        data.expectedMonth,
        data.expectedYear,
        data.bankCol || "",
        data.expectedDay || ""
      ]]);
      return "Spesa programmata modificata.";
    }
  }
  throw new Error("Spesa programmata non trovata.");
}
