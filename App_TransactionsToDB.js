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
  // FIX (1.10): LockService per evitare race condition. Senza lock, un doppio tap su mobile
  // o due chiamate concorrenti possono leggere la STESSA "prima riga vuota" e sovrascriversi
  // a vicenda, perdendo una transazione senza alcun errore.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return "Error: Sistema occupato, riprova tra qualche secondo.";
  }

  try {
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
  } finally {
    lock.releaseLock(); // FIX (1.10): rilascia sempre il lock, anche sui return di errore
  }
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
  
  // FIX (3.2): WRITE-BEFORE-DELETE. Si individua la riga SENZA eliminarla; si scrive prima
  // lo storico e il rimborso, e SOLO se entrambe le scritture vanno a buon fine si elimina
  // la riga da Active_Credits. Cosi' un errore di scrittura non fa sparire il credito.
  let rowToDelete = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      rowData = data[i];   // Save original data
      rowToDelete = i + 1; // Riga reale sul foglio (NON ancora eliminata)
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Credito non trovato o già saldato.");

  // 2. Save to History (Settled_Credits) -- PRIMA della delete
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

  // 4. Record the Refund transaction -- anch'essa PRIMA della delete
  const refundResult = addTransaction({
    type: 'Refund',
    category: category,
    details: "Settled from: " + note,
    amounts: amountsObj
  });

  // 5. Solo ora che storico + rimborso sono stati scritti, rimuovi da Active_Credits
  creditSheet.deleteRow(rowToDelete);

  return refundResult;
}

/**
 * Segna un credito come perso (Lost/Bad Debt): lo sposta nel foglio Settled_Credits
 * senza creare una transazione di rimborso (nessuna spesa/entrata viene scritta).
 */
function markCreditAsLost(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creditSheet = ss.getSheetByName("Active_Credits");
  const settledSheet = ss.getSheetByName("Settled_Credits");
  
  if (!creditSheet) throw new Error("Foglio Active_Credits mancante.");
  if (!settledSheet) throw new Error("Crea prima il foglio 'Settled_Credits'!");

  const data = creditSheet.getDataRange().getValues();
  let found = false;
  let rowData = [];
  
  let rowToDelete = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      rowData = data[i];   // Save original data
      rowToDelete = i + 1; // Riga reale sul foglio
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Credito non trovato o già gestito.");

  // 1. Save to History (Settled_Credits) with a special status "Lost"
  const settleDate = new Date();
  settledSheet.appendRow([
    rowData[0], // Original ID
    rowData[1], // Original Date
    rowData[2], // Who
    rowData[3], // Category
    rowData[4], // Note
    rowData[5], // Amount
    settleDate, // Settled Date
    "Lost"      // Bank Col equivalent indicating Lost/Bad Debt
  ]);

  // 2. Elimina la riga da Active_Credits senza creare il refund
  creditSheet.deleteRow(rowToDelete);

  return { success: true, message: "Credito segnato come perso." };
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
 * Records a standalone debt and auto-creates the monthly fixed expense.
 * Supports standard amortization and Balloon payments (Maxirata).
 */
function addStandaloneDebt(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbDebts = ss.getSheetByName("DB_Debts");
  if (!dbDebts) throw new Error("Sheet DB_Debts not found.");

  const debtId = "DBT-" + new Date().getTime().toString().slice(-6);
  
  let principal = parseFloat(data.amount);
  let balloon = parseFloat(data.balloon) || 0;
  let rateDec = parseFloat(data.rate) / 100; 
  let years = parseFloat(data.years);
  let months = Math.round(years * 12);
  
  let monthlyRate = rateDec / 12;
  let monthlyPayment = 0;
  
  if (monthlyRate === 0) {
      monthlyPayment = (principal - balloon) / months;
  } else {
      monthlyPayment = ((principal * Math.pow(1 + monthlyRate, months)) - balloon) * monthlyRate / (Math.pow(1 + monthlyRate, months) - 1);
  }

  let start = new Date(data.date);
  let end = new Date(start);
  end.setMonth(end.getMonth() + months);

  dbDebts.appendRow([
    debtId, data.name, data.date, principal, rateDec, years,
    Number(monthlyPayment.toFixed(2)), "Active", balloon > 0 ? balloon : ""
  ]);

  try {
      addFixedExpenseToSheet({
        cat: "Fees Taxes", // <--- FIX CATEGORIA
        note: "Rata " + data.name,
        amt: Number(monthlyPayment.toFixed(2)),
        startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        payDay: start.getDate(),
        bankCol: data.bankCol, 
        isSplit: false
      });
  } catch (e) { Logger.log("Errore spese fisse: " + e.message); }

  return "Debt recorded successfully.";
}

/**
 * Recupera tutti i debiti con stato "Active" per la tendina del Wizard.
 */
function getActiveDebts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_Debts");
  if(!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let active = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][7] === "Active") {
      active.push({
        id: data[i][0],
        name: data[i][1],
        startDate: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'dd/MM/yyyy') : data[i][2]
      });
    }
  }
  return active;
}

/**
 * Esegue la logica completa di Rinegoziazione o Surroga:
 * 1. Calcola debito residuo.
 * 2. Archivia vecchio debito.
 * 3. Crea nuovo debito e lo ricollega agli Asset.
 * 4. Stoppa la vecchia rata mensile e ne crea una nuova.
 */
function renegotiateDebtBackend(oldDebtId, newRatePct, newYears, newDateStr, bankCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbDebts = ss.getSheetByName("DB_Debts");
  const dbAssets = ss.getSheetByName("DB_RealAssets");
  const dbFixed = ss.getSheetByName("Config_FixedExpenses");
  
  const debtsData = dbDebts.getDataRange().getValues();
  let oldDebtRowIdx = -1;
  let oldDebt = null;
  
  // Trova il debito da chiudere
  for(let i = 1; i < debtsData.length; i++) {
    if(debtsData[i][0] === oldDebtId) {
      oldDebtRowIdx = i + 1;
      oldDebt = {
         name: debtsData[i][1],
         startDate: new Date(debtsData[i][2]),
         principal: parseFloat(debtsData[i][3]),
         rateDec: parseFloat(debtsData[i][4]), // Questo è già es. 0.025
         years: parseFloat(debtsData[i][5])
      };
      break;
    }
  }
  if(!oldDebt) throw new Error("Debito originale non trovato nel DB.");

  // --- CALCOLO DEBITO RESIDUO (Formula Inversa Ammortamento Francese) ---
  const newDate = new Date(newDateStr);
  let monthsPassed = (newDate.getFullYear() - oldDebt.startDate.getFullYear()) * 12 + (newDate.getMonth() - oldDebt.startDate.getMonth());
  
  let oldMonthlyRate = oldDebt.rateDec / 12;
  let oldTotalMonths = oldDebt.years * 12;
  let outstandingAmount = oldDebt.principal; 
  
  if(monthsPassed > 0 && monthsPassed < oldTotalMonths) {
     // Calcolo Rata originaria
     let pmt = oldDebt.principal * (oldMonthlyRate * Math.pow(1 + oldMonthlyRate, oldTotalMonths)) / (Math.pow(1 + oldMonthlyRate, oldTotalMonths) - 1);
     if(isNaN(pmt)) pmt = oldDebt.principal / oldTotalMonths; // Fallback tasso zero
     
     // Attualizzazione delle rate rimanenti
     let monthsRemaining = oldTotalMonths - monthsPassed;
     outstandingAmount = pmt * (1 - Math.pow(1 + oldMonthlyRate, -monthsRemaining)) / oldMonthlyRate;
  } else if (monthsPassed >= oldTotalMonths) {
     throw new Error("Impossibile rinegoziare: questo debito risulta già estinto nella data indicata.");
  }

  // 1. Chiudi vecchio debito
  dbDebts.getRange(oldDebtRowIdx, 8).setValue("Renegotiated");

  // 2. Crea nuovo debito
  const newDebtId = "DBT-" + new Date().getTime().toString().slice(-6);
  let newRateDec = parseFloat(newRatePct) / 100; // Come sempre formattiamo il tasso
  let newMonths = Math.round(parseFloat(newYears) * 12);
  let newMonthlyRate = newRateDec / 12;
  
  let newPmt = outstandingAmount * (newMonthlyRate * Math.pow(1 + newMonthlyRate, newMonths)) / (Math.pow(1 + newMonthlyRate, newMonths) - 1);
  if(isNaN(newPmt) || newMonthlyRate === 0) newPmt = outstandingAmount / newMonths;

  dbDebts.appendRow([
    newDebtId,
    oldDebt.name + " (Rin.)",
    newDateStr,
    outstandingAmount, // Il nuovo capitale erogato è il residuo precedente!
    newRateDec,
    newYears,
    Number(newPmt.toFixed(2)),
    "Active"
  ]);

  // 3. Aggiorna DB_RealAssets (Se il debito è legato a una casa, sostituisce il link)
  if (dbAssets) {
     let assetsData = dbAssets.getDataRange().getValues();
     for(let i = 1; i < assetsData.length; i++) {
        if(assetsData[i][10] === oldDebtId) { // Linked_Debt_ID è colonna K
           dbAssets.getRange(i + 1, 11).setValue(newDebtId);
        }
     }
  }

  // 4. Aggiorna Spese Fisse (Config_FixedExpenses)
  if (dbFixed) {
     let fixedData = dbFixed.getDataRange().getValues();
     // Blocca l'addebito della vecchia rata sovrascrivendo l'EndDate con la data di rinegoziazione
     for(let i = 1; i < fixedData.length; i++) {
        let note = String(fixedData[i][2]);
        if(note.includes(oldDebt.name) && fixedData[i][8] !== true) { // Cerca la nota e assicurati non sia uno split
           dbFixed.getRange(i + 1, 8).setValue(newDateStr); // Colonna H: endDate
        }
     }
     
     // Calcola la scadenza del nuovo accordo
     let endNewDate = new Date(newDate);
     endNewDate.setMonth(endNewDate.getMonth() + newMonths);
     
     addFixedExpenseToSheet({
        cat: oldDebt.name.includes("Mutuo") ? "Mutuo Immobile" : "Debiti/Prestiti",
        note: "Rata " + oldDebt.name + " (Rin.)",
        amt: Number(newPmt.toFixed(2)),
        startDate: Utilities.formatDate(newDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        endDate: Utilities.formatDate(endNewDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        payDay: newDate.getDate(),
        bankCol: bankCol,
        isSplit: false
     });
  }

  return "Rinegoziazione completata con successo!\nNuovo capitale ricalcolato: € " + outstandingAmount.toFixed(2) + "\nNuova rata: € " + newPmt.toFixed(2);
}

/**
 * Recupera gli asset attivi per la UI.
 */
function getActiveRealAssets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_RealAssets");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let active = [];
  // Evitiamo la riga di intestazione partendo da i = 1
  for(let i = 1; i < data.length; i++) {
    if(data[i][11] === "Active") active.push({ id: data[i][0], type: data[i][1], name: data[i][2] });
  }
  return active;
}

/**
 * BACKEND: Gestisce Estinzioni Debiti (Totali o Parziali)
 */
function repayDebtBackend(payload) {
  const { id, type, amt, bank } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbDebts = ss.getSheetByName("DB_Debts");
  const dbFixed = ss.getSheetByName("Config_FixedExpenses");
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const debtsData = dbDebts.getDataRange().getValues();
  let debtRow = -1; let debt = null;
  
  for(let i = 1; i < debtsData.length; i++) {
    if(debtsData[i][0] === id) {
      debtRow = i + 1;
      debt = { 
          name: debtsData[i][1], start: new Date(debtsData[i][2]), principal: parseFloat(debtsData[i][3]), 
          rate: parseFloat(debtsData[i][4]), years: parseFloat(debtsData[i][5]), balloon: parseFloat(debtsData[i][8]) || 0 
      }; break;
    }
  }
  if(!debt) throw new Error("Debito non trovato.");

  let monthsPassed = (new Date().getFullYear() - debt.start.getFullYear()) * 12 + (new Date().getMonth() - debt.start.getMonth());
  let mRate = debt.rate / 12; let totMonths = debt.years * 12;
  let outstanding = debt.principal;
  
  if(monthsPassed > 0 && monthsPassed < totMonths) {
     let pmt = 0;
     if (mRate === 0) pmt = (debt.principal - debt.balloon) / totMonths;
     else pmt = ((debt.principal * Math.pow(1 + mRate, totMonths)) - debt.balloon) * mRate / (Math.pow(1 + mRate, totMonths) - 1);
     
     let remMonths = totMonths - monthsPassed;
     if (mRate === 0) outstanding = (pmt * remMonths) + debt.balloon;
     else outstanding = (pmt * (1 - Math.pow(1 + mRate, -remMonths)) / mRate) + (debt.balloon * Math.pow(1 + mRate, -remMonths));
  }

  let amountPaid = type === 'Total' ? outstanding : parseFloat(amt);

  // Registra l'uscita di cassa
  let amountsObj = {}; amountsObj[bank] = -Math.abs(amountPaid); 
  addTransaction({ 
      type: 'Expense', 
      category: "Fees Taxes", // <--- FIX CATEGORIA (Invece di Debiti/Prestiti)
      details: `Estinzione ${type} - ${debt.name}`, 
      amounts: amountsObj 
  });

  if(type === 'Total') {
     dbDebts.getRange(debtRow, 8).setValue("Paid_Off");
     if(dbFixed) {
        let fData = dbFixed.getDataRange().getValues();
        for(let i=1; i<fData.length; i++) {
           if(String(fData[i][2]).includes(debt.name) && fData[i][8] !== true) dbFixed.getRange(i+1, 8).setValue(todayStr); 
        }
     }
     return "Estinzione totale completata. Pagati €" + amountPaid.toFixed(2);
  } else {
     let newPrincipal = outstanding - amountPaid;
     if(newPrincipal <= 0) throw new Error("Importo superiore al residuo. Usa l'estinzione totale.");
     
     dbDebts.getRange(debtRow, 8).setValue("Partially_Paid"); 
     const newId = "DBT-" + new Date().getTime().toString().slice(-6);
     
     let remMonths = totMonths - monthsPassed;
     let newPmt = 0;
     if(mRate === 0) newPmt = (newPrincipal - debt.balloon) / remMonths;
     else newPmt = ((newPrincipal * Math.pow(1 + mRate, remMonths)) - debt.balloon) * mRate / (Math.pow(1 + mRate, remMonths) - 1);

     dbDebts.appendRow([newId, debt.name + " (Ridotto)", todayStr, newPrincipal, debt.rate, (remMonths/12), Number(newPmt.toFixed(2)), "Active", debt.balloon > 0 ? debt.balloon : ""]);
     
     const dbAssets = ss.getSheetByName("DB_RealAssets");
     if(dbAssets) {
        let aData = dbAssets.getDataRange().getValues();
        for(let i=1; i<aData.length; i++) if(aData[i][10] === id) dbAssets.getRange(i+1, 11).setValue(newId);
     }

     if(dbFixed) {
        let fData = dbFixed.getDataRange().getValues();
        for(let i=1; i<fData.length; i++) {
           if(String(fData[i][2]).includes(debt.name) && fData[i][8] !== true) dbFixed.getRange(i+1, 8).setValue(todayStr);
        }
        
        let start = new Date();
        let end = new Date();
        end.setMonth(end.getMonth() + remMonths);
        addFixedExpenseToSheet({
          cat: debt.name.includes("Mutuo") ? "Housing" : "Fees Taxes", // <--- FIX CATEGORIE (Dipende se è mutuo o debito normale)
          note: "Rata " + debt.name + " (Ridotto)",
          amt: Number(newPmt.toFixed(2)),
          startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          payDay: start.getDate(),
          bankCol: bank, 
          isSplit: false
        });
     }

     return `Estinzione parziale completata. Nuovo residuo: €${newPrincipal.toFixed(2)}`;
  }
}
