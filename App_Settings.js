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
    Investment: [],
    Refund: [] // <-- AGGIUNTO PER SUPPORTO RIMBORSI
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
  // ... (Il resto della funzione rimane identico) ...
  const expSheet = ss.getSheetByName("Config_FixedExpenses");
  let fixedExpensesData = [];

  if (expSheet) {
    const expRows = expSheet.getDataRange().getValues();
    for (let i = 1; i < expRows.length; i++) {
      let row = expRows[i];
      if (row[0] === "") continue; 

      let rawAmt = String(row[3]).replace('€', '').replace(',', '.').trim();
      let amount = parseFloat(rawAmt) || 0;

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

      if (sub.isSplit && row[9]) {
        let splits = [];
        let splitParts = String(row[9]).split('|'); 
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

/**
 * Elimina una Spesa Fissa dal foglio Settings in base all'ID
 */
function deleteFixedExpenseFromSheet(expenseId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config_FixedExpenses");
  if (!sheet) throw new Error("Foglio Config_FixedExpenses mancante");

  const data = sheet.getDataRange().getValues();
  // Partiamo da 1 per saltare l'intestazione
  for (let i = 1; i < data.length; i++) { 
    if (String(data[i][0]) === String(expenseId)) {
      const name = data[i][2]; // Il nome/nota si trova in colonna C (indice 2)
      sheet.deleteRow(i + 1);  // i+1 perché Apps Script è basato su indici 1 per le righe
      return `Spesa "${name}" eliminata!`;
    }
  }
  throw new Error("Spesa fissa non trovata nel database.");
}

/**
 * Elimina una Categoria dal foglio Settings
 * Se viene eliminata una Expense, elimina anche la rispettiva categoria Refund.
 */
function deleteCategoryFromSheet(type, categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config_Category");
  if (!sheet) throw new Error("Foglio Config_Category mancante");

  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;
  
  // Ciclo al contrario per permettere l'eliminazione di più righe in sicurezza
  for (let i = data.length - 1; i >= 1; i--) { 
    const rowType = data[i][0];
    const rowCat = data[i][1];
    
    // Elimina la riga originariamente selezionata dall'utente
    if (rowType === type && rowCat === categoryName) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
    // Elimina in automatico anche il Refund collegato se si sta eliminando una Expense
    else if (type === "Expense" && rowType === "Refund" && rowCat === categoryName) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  
  if (deletedCount > 0) {
      return `Categoria "${categoryName}" eliminata con successo!`;
  }
  throw new Error("Categoria non trovata nel database.");
}

/**
 * Aggiunge una nuova categoria. 
 * Se è di tipo 'Expense', ne crea automaticamente una identica di tipo 'Refund'.
 */
function addCategoryToSheet(type, categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config_Category");
  if (!sheet) throw new Error("Foglio Config_Category mancante");

  // Aggiunge la categoria principale
  sheet.appendRow([type, categoryName]);

  // Seleziona se creare anche la controparte "Refund"
  if (type === "Expense") {
      sheet.appendRow(["Refund", categoryName]);
  }
  
  return "Categoria aggiunta!";
}