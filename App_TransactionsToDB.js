/**
 * Main function to record standard transactions (Expenses/Income).
 * Features an auto-sync mechanism that replicates 'Investment' entries into the 
 * respective History sheets (Stocks/Crypto) as 'Cash Deposits', linking them via a unique ID.
 * * @param {Object} data - Transaction data object.
 * @param {string} data.type - Transaction type (Expense, Income, Investment).
 * @param {string} data.category - Category (e.g., "Alimentazione", "Azioni").
 * @param {string} data.details - Description or notes.
 * @param {Object} data.amounts - Key-value pair of { columnIndex: amount }.
 * @return {string} Success message.
 */
function addTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Dynamically determine the target sheet based on the current year
  const dateVal = new Date();
  const currentYear = dateVal.getFullYear();
  const sheetName = "Expenses Tracker " + currentYear;
  
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet not found (" + sheetName + ")";

  const startRow = 20;
  
  // Logic to find the first truly empty row
  const colB = sheet.getRange(startRow, 2, Math.max(1, sheet.getLastRow() - startRow + 1), 1).getValues();
  let newRow = startRow;
  for (let i = 0; i < colB.length; i++) {
    if (colB[i][0] === "" || colB[i][0] === null) {
      newRow = startRow + i;
      break;
    }
    if (i === colB.length - 1) newRow = startRow + i + 1;
  }

  // Find "Controllo Automatismi" column dynamically
  const headerRowTx = 2;
  const maxCols = sheet.getLastColumn();
  const headersTx = sheet.getRange(headerRowTx, 1, 1, maxCols).getValues()[0];
  const syncColIndexTx = headersTx.indexOf("Controllo Automatismi") + 1;
  
  // PREPARE OPTIMIZED BATCH ARRAY FOR EXPENSES
  const numCols = Math.max(maxCols, syncColIndexTx);
  let txRowData = new Array(numCols).fill("");

  txRowData[0] = dateVal;
  txRowData[1] = data.type;
  txRowData[2] = data.category;
  txRowData[3] = data.details;

  let totalInvestAmount = 0;
  for (let col in data.amounts) {
    let colIdx = parseInt(col);
    let val = parseFloat(data.amounts[col]);
    
    if (colIdx > 0) {
      txRowData[colIdx - 1] = val; // -1 because Array is 0-indexed
    }
    
    // IMPORTANT: Verify these column numbers match your new bank columns!
    if (colIdx >= 5 && colIdx <= 9) {
      totalInvestAmount += Math.abs(val || 0);
    }
  }

  let generatedTxId = null;

  // --- 2. IMMEDIATE SYNCHRONIZATION (AUTO-LINK) ---
  if (data.type === "Investment") {
    let destSheetName = null;
    let isCrypto = false;

    if (data.category === "Stocks") destSheetName = "History B/S Stocks";
    if (data.category === "Crypto") { destSheetName = "History B/S Crypto"; isCrypto = true; }

    if (destSheetName) {
      const destSheet = ss.getSheetByName(destSheetName);
      if (destSheet) {
        generatedTxId = "ID_" + new Date().getTime() + "_" + newRow;

        const lastHistRow = destSheet.getLastRow();
        let histRow = 1;
        const histDates = destSheet.getRange("A1:A" + (lastHistRow + 1)).getValues();
        for (let j = histDates.length - 1; j >= 0; j--) {
          if (histDates[j][0] !== "" && histDates[j][0] != null) {
            histRow = j + 2; break;
          }
        }
        if (lastHistRow === 0) histRow = 1;

        // OPTIMIZED BATCH WRITE TO HISTORY
        let histRowData = new Array(13).fill("");
        histRowData[0] = dateVal;
        histRowData[1] = "Cash";
        histRowData[2] = "Deposit";
        histRowData[3] = 1;
        histRowData[4] = "x";
        histRowData[5] = totalInvestAmount;
        
        if (isCrypto) {
          histRowData[7] = totalInvestAmount; // Col H (8)
        }
        histRowData[12] = generatedTxId; // Col M (13)

        // Write History in one call
        destSheet.getRange(histRow, 1, 1, 13).setValues([histRowData]);
      }
    }
  }

  // --- 3. SPLIT WITH FRIENDS LOGIC ---
  let splitDataToAppend = [];
  if (data.splitData && data.splitData.length > 0) {
    if (!generatedTxId) {
      generatedTxId = "TX_" + new Date().getTime() + "_" + newRow;
    }
    
    data.splitData.forEach(friend => {
      const creditId = "CR_" + new Date().getTime() + "_" + Math.floor(Math.random()*1000);
      splitDataToAppend.push([creditId, dateVal, friend.who, data.category, data.details, friend.amount, generatedTxId]);
    });
  }

  // Add the Sync ID to the Expenses Array if it was generated
  if (generatedTxId && syncColIndexTx > 0) {
    txRowData[syncColIndexTx - 1] = generatedTxId;
  }

  // WRITE EXPENSES ROW IN ONE SINGLE CALL (Massive speedup)
  sheet.getRange(newRow, 1, 1, numCols).setValues([txRowData]);

  // Write Active Credits if needed
  if (splitDataToAppend.length > 0) {
    const creditSheet = ss.getSheetByName("Active_Credits");
    if (creditSheet) {
      const creditStartRow = creditSheet.getLastRow() + 1;
      creditSheet.getRange(creditStartRow, 1, splitDataToAppend.length, 7).setValues(splitDataToAppend);
    }
  }

  //if (generatedTxId) SpreadsheetApp.flush();
  
  return "Saved Successfully";
}
/**
 * Recupera la lista dei crediti attivi dal foglio
 */
function getActiveCredits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Active_Credits");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let credits = [];
  
  for (let i = 1; i < data.length; i++) {
    if(data[i][0]) { 
      credits.push({
        id: data[i][0],
        date: data[i][1] ? Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        who: data[i][2],
        category: data[i][3],
        note: data[i][4],
        amount: parseFloat(data[i][5]) || 0,
        linkedTxId: data[i][6] || "" // Recupera l'ID collegato dalla colonna 7 (G)
      });
    }
  }
  return credits;
}

/**
 * Salda il debito: lo sposta nel foglio Settled_Credits e crea un "Refund"
 */
function settleActiveCredit(id, amount, category, note, bankCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creditSheet = ss.getSheetByName("Active_Credits");
  const settledSheet = ss.getSheetByName("Settled_Credits");
  
  if (!creditSheet) throw new Error("Foglio Active_Credits mancante.");
  if (!settledSheet) throw new Error("Crea prima il foglio 'Settled_Credits'!");

  const data = creditSheet.getDataRange().getValues();
  let found = false;
  let rowData = [];
  
  // 1. Find, copy, and delete the row from Active_Credits
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      rowData = data[i]; // Save original data
      creditSheet.deleteRow(i + 1); 
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Credito non trovato o già saldato.");

  // 2. Save to History (Settled_Credits)
  const settleDate = new Date();
  settledSheet.appendRow([
    rowData[0], // Original ID
    rowData[1], // Original Date
    rowData[2], // Who
    rowData[3], // Category
    rowData[4], // Note
    rowData[5], // Amount
    settleDate, // Settled Date
    bankCol     // Bank Col
  ]);

  // 3. Create amounts object for the Refund
  let amountsObj = {};
  amountsObj[bankCol] = amount; 

  // 4. Record the Refund transaction
  return addTransaction({
    type: 'Refund',
    category: category,
    details: "Settled from: " + note,
    amounts: amountsObj
  });
}

/**
 * Salda tutti i debiti di una persona specifica in blocco.
 * Conserva la precisione delle categorie dividendo il rimborso su più transazioni.
 */
function settleGroupedCredits(normalizedName, bankCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creditSheet = ss.getSheetByName("Active_Credits");
  const settledSheet = ss.getSheetByName("Settled_Credits");

  if (!creditSheet || !settledSheet) throw new Error("Fogli crediti mancanti.");

  const data = creditSheet.getDataRange().getValues();
  const settleDate = new Date();
  let itemsToSettle = [];
  
  // 1. Trova ed estrai TUTTE le righe associate a quella persona (partendo dal basso)
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]).trim().toUpperCase() === normalizedName) {
      itemsToSettle.push(data[i]);
      creditSheet.deleteRow(i + 1);
    }
  }

  if (itemsToSettle.length === 0) throw new Error("Nessun credito trovato per questa persona.");

 let categoryTotals = {};
  let settledRowsToAppend = [];
  
  // 2. Prepara l'array per Settled_Credits e raggruppa le somme (OPTIMIZED)
  itemsToSettle.forEach(row => {
    settledRowsToAppend.push([
      row[0], row[1], row[2], row[3], row[4], row[5], settleDate, bankCol
    ]);

    let cat = row[3];
    let amt = parseFloat(row[5]) || 0;
    if(!categoryTotals[cat]) categoryTotals[cat] = 0;
    categoryTotals[cat] += amt;
  });

  // Scrive nello storico Settled_Credits in una singola operazione
  if (settledRowsToAppend.length > 0) {
    const lastSettledRow = settledSheet.getLastRow();
    settledSheet.getRange(lastSettledRow + 1, 1, settledRowsToAppend.length, 8).setValues(settledRowsToAppend);
  }

  // 3. Crea automaticamente un Refund separato per ogni categoria coinvolta
  let resultMsg = "Saldato";
  let personName = itemsToSettle[0][2]; // Nome originale per la nota
  
  for (let cat in categoryTotals) {
     let amountsObj = {};
     amountsObj[bankCol] = categoryTotals[cat]; // Associa l'importo totale della categoria al conto scelto

     // Sfruttiamo la tua funzione addTransaction per fare il lavoro sporco
     resultMsg = addTransaction({
       type: 'Refund',
       category: cat,
       details: `Bulk settlement from: ${personName}`, 
       amounts: amountsObj
     });
  }
  
  return resultMsg;
}
/**
 * Records a standalone debt (Loan, Personal Financing) in the DB_Debts sheet.
 */
function addStandaloneDebt(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbDebts = ss.getSheetByName("DB_Debts");
  
  if (!dbDebts) throw new Error("Sheet DB_Debts not found.");

  const debtId = "DBT-" + new Date().getTime().toString().slice(-6);
  
  dbDebts.appendRow([
    debtId,
    data.name,
    data.date,
    parseFloat(data.amount),
    parseFloat(data.rate),
    parseFloat(data.years),
    "", // Monthly payment
    "Active"
  ]);

  return "Debt recorded successfully.";
}