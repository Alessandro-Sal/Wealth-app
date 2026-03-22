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
  // Target Sheet: Expenses Tracker 2026 (Hardcoded year for now)
  const sheet = ss.getSheetByName("Expenses Tracker 2026");
  if (!sheet) return "Error: Sheet not found";

  const startRow = 20;
  
  // Logic to find the first truly empty row in the tracker
  const colB = sheet.getRange(startRow, 2, Math.max(1, sheet.getLastRow() - startRow + 1), 1).getValues();
  let newRow = startRow;
  for (let i = 0; i < colB.length; i++) {
    if (colB[i][0] === "" || colB[i][0] === null) {
      newRow = startRow + i;
      break;
    }
    if (i === colB.length - 1) newRow = startRow + i + 1;
  }

  // 1. WRITE TO EXPENSES TRACKER (Standard Operation)
  const dateVal = new Date();
  sheet.getRange(newRow, 1).setValue(dateVal);
  sheet.getRange(newRow, 2).setValue(data.type);
  sheet.getRange(newRow, 3).setValue(data.category);
  sheet.getRange(newRow, 4).setValue(data.details);

  // Write amounts and calculate total for potential investment sync
  let totalInvestAmount = 0;
  for (let col in data.amounts) {
    let val = parseFloat(data.amounts[col]);
    sheet.getRange(newRow, parseInt(col)).setValue(val);
    
    // Sum absolute values if columns are between 5 (E) and 9 (I) - Typical Bank Columns
    if (parseInt(col) >= 5 && parseInt(col) <= 9) {
      totalInvestAmount += Math.abs(val || 0);
    }
  }

  // --- 2. IMMEDIATE SYNCHRONIZATION (AUTO-LINK) ---
  // If Type is "Investment", immediately copy to History as a "Deposit"
  if (data.type === "Investment") {
    
    let destSheetName = null;
    let isCrypto = false;

    // Check Category to determine destination (Strings must match Dropdown values)
    if (data.category === "Azioni") destSheetName = "History B/S Stocks";
    if (data.category === "Crypto") { destSheetName = "History B/S Crypto"; isCrypto = true; }

    if (destSheetName) {
      const destSheet = ss.getSheetByName(destSheetName);
      if (destSheet) {
        // Generate Unique ID based on timestamp and row
        const newId = "ID_" + new Date().getTime() + "_" + newRow;

        // Find empty row in History sheet
        const lastHistRow = destSheet.getLastRow();
        let histRow = 1;
        
        // Search for first free row by checking Column A (Date) backwards
        const histDates = destSheet.getRange("A1:A" + (lastHistRow + 1)).getValues();
        for (let j = histDates.length - 1; j >= 0; j--) {
          if (histDates[j][0] !== "" && histDates[j][0] != null) {
            histRow = j + 2;
            break;
          }
        }
        if (lastHistRow === 0) histRow = 1;

        // Write to History: A=Date, B=Ticker(Cash), C=Action(Deposit), D=Qty(1), E=Class(x), F=Amount
        destSheet.getRange(histRow, 1, 1, 6).setValues([[dateVal, "Cash", "Deposit", 1, "x", totalInvestAmount]]);
        
        // If Crypto, write amount to Col H (8) as well (specific formatting for Crypto sheet)
        if (isCrypto) {
          destSheet.getRange(histRow, 8).setValue(totalInvestAmount);
        }

        // Write ID to History (Col M = 13)
        destSheet.getRange(histRow, 13).setValue(newId);

        // WRITE ID TO EXPENSES (Col AI = 35) - Links the two rows
        sheet.getRange(newRow, 35).setValue(newId);
        
        // Force save to ensure data integrity across sheets
        SpreadsheetApp.flush();
      }
    }
  }
// --- 3. SPLIT WITH FRIENDS LOGIC ---
  if (data.splitData && data.splitData.length > 0) {
    const creditSheet = ss.getSheetByName("Active_Credits");
    if (creditSheet) {
      // Genera un ID di transazione unico se non esiste (lo useremo per collegare credito e spesa)
      const txId = "TX_" + new Date().getTime() + "_" + newRow;
      // Salva l'ID nella riga della spesa in Expenses Tracker (usiamo la colonna 35 / AI come fai per gli investimenti)
      sheet.getRange(newRow, 35).setValue(txId);

      data.splitData.forEach(friend => {
        const creditId = "CR_" + new Date().getTime() + "_" + Math.floor(Math.random()*1000);
        
        creditSheet.appendRow([
          creditId,
          dateVal,
          friend.who,
          data.category,
          data.details,
          friend.amount,
          txId // Salviamo l'ID della transazione originale nella colonna 7 (G) di Active_Credits
        ]);
      });
    }
  }
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
  
  // 2. Scrivi tutto nello storico (Settled_Credits) e raggruppa le somme per Categoria
  itemsToSettle.forEach(row => {
    settledSheet.appendRow([
      row[0], row[1], row[2], row[3], row[4], row[5], settleDate, bankCol
    ]);

    let cat = row[3];
    let amt = parseFloat(row[5]) || 0;
    if(!categoryTotals[cat]) categoryTotals[cat] = 0;
    categoryTotals[cat] += amt;
  });

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