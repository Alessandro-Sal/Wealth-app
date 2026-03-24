/**
 * Adds a new investment transaction to the History sheets or Real Assets DB.
 * If the action is a "Withdrawal" (Cash Out), it automatically syncs the entry 
 * to the current year's "Expenses Tracker" as an Investment expense/transfer.
 * Generates a unique Transaction ID to allow synchronized deletion later.
 * * @param {Object} data - The transaction data object.
 * @param {string} data.type - "Crypto", "Stocks", "RealEstate", or "Bond".
 * @param {string} data.action - "Deposit", "Withdrawal", "Buy", "Sell".
 * @param {string} data.ticker - The asset symbol (e.g., "BTC", "AAPL").
 * @param {number} data.qty - Quantity of the asset.
 * @param {number} data.costEur - Total cost in EUR.
 * @param {number} data.costUsd - Total cost in USD.
 * @param {string} [data.bankCol] - The column index of the bank used (required for withdrawals).
 * @param {string} [data.note] - Optional user notes.
 * @return {string} Status message.
 */
function addInvestTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- NEW: REAL ASSETS & BONDS (Write to DB_RealAssets & DB_Debts) ---
  if (data.type === "RealEstate" || data.type === "Bond") {
    const dbAssets = ss.getSheetByName("DB_RealAssets");
    if (!dbAssets) return "Error: DB_RealAssets sheet missing.";
    
    let assetId = (data.type === "RealEstate" ? "RE-" : "BND-") + new Date().getTime().toString().slice(-6);
    let linkedDebtId = "";
    
    // Se è un immobile col Mutuo, crea il Debito e la Spesa Fissa
    if (data.type === "RealEstate" && parseFloat(data.loanAmt) > 0) {
      const dbDebts = ss.getSheetByName("DB_Debts");
      if (dbDebts) {
        linkedDebtId = "DBT-" + new Date().getTime().toString().slice(-6);
        
        let principal = parseFloat(data.loanAmt);
        let rateDec = parseFloat(data.rate) / 100; // FIX 250%
        let years = parseFloat(data.years);
        let months = Math.round(years * 12);
        
        let monthlyRate = rateDec / 12;
        let monthlyPayment = principal * (monthlyRate * Math.pow(1 + monthlyRate, months)) / (Math.pow(1 + monthlyRate, months) - 1);
        if (isNaN(monthlyPayment) || monthlyRate === 0) monthlyPayment = principal / months;

        let start = new Date(data.startDate || data.date);
        let end = new Date(start);
        end.setMonth(end.getMonth() + months);

        dbDebts.appendRow([
          linkedDebtId, "Mutuo Ipotecario", data.startDate || data.date, 
          principal, rateDec, years, Number(monthlyPayment.toFixed(2)), "Active"
        ]);

        // Auto-add alla Dashboard delle spese fisse
        try {
            addFixedExpenseToSheet({
              cat: "Mutuo Immobile",
              note: "Rata " + data.name,
              amt: Number(monthlyPayment.toFixed(2)),
              startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd"),
              endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd"),
              payDay: start.getDate(),
              bankCol: data.bankCol, // <--- NUOVO: Passiamo la banca al DB
              isSplit: false
            });
        } catch(e) {}
      }
    }

    let rowData = new Array(14).fill("");
    rowData[0] = assetId;
    rowData[1] = data.type === "RealEstate" ? "Real Estate" : (data.tax === "true" ? "Government Bond" : "Corporate Bond");
    rowData[2] = data.name;
    rowData[3] = data.type === "Bond" ? data.isin : ""; 
    rowData[4] = data.startDate || data.date; 
    rowData[11] = "Active"; 

    if (data.type === "RealEstate") {
      let marketVal = parseFloat(data.marketVal);
      let loanAmt = parseFloat(data.loanAmt) || 0;
      
      rowData[5] = marketVal; // Purchase Price
      rowData[10] = linkedDebtId; 

      // --- NUOVO: Sottrae l'anticipo versato dal conto in banca ---
      let upfrontCash = marketVal - loanAmt;
      if (upfrontCash > 0 && data.bankCol) {
          let amountsObj = {};
          amountsObj[data.bankCol] = upfrontCash;
          try {
              // Sfruttiamo il motore transazioni per registrare l'uscita
              addTransaction({
                  type: 'Expense',
                  category: 'Real Estate',
                  details: 'Acquisto / Anticipo: ' + data.name,
                  amounts: amountsObj
              });
          } catch(e) { Logger.log("Errore registrazione anticipo cassa: " + e.message); }
      }

    } else if (data.type === "Bond") {
      let nominal = parseFloat(data.nominal);
      let price = parseFloat(data.price);
      rowData[5] = (nominal * price) / 100; 
      rowData[6] = nominal; 
      rowData[7] = parseFloat(data.coupon) / 100; // FIX 250% anche per le cedole
      rowData[9] = data.tax === "true" ? "WhiteList" : "Standard"; 
      rowData[12] = data.isin; 
    }
    
    dbAssets.appendRow(rowData);
    return data.type === "RealEstate" ? "Real Estate & Mortgage Added!" : "Bond Added!";
  }

  
  // Generate a unique ID to link History and Expenses entries (crucial for synchronized deletion)
  const transactionId = "ID_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);

  // --- PART 1: INVESTMENT HISTORY (History B/S) ---
  const isCrypto = data.type === "Crypto";
  const sheetName = isCrypto ? "History B/S Crypto" : "History B/S Stocks";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet missing (" + sheetName + ")";

  // Find first empty row in History (Starting from Row 3)
  const lastRow = sheet.getLastRow();
  let newRow = lastRow + 1;
  const startRowHist = 3;
  
  if (lastRow >= startRowHist) {
    // Scan Column A to find the first actual empty cell
    const colA = sheet.getRange(startRowHist, 1, lastRow - startRowHist + 2, 1).getValues();
    for (let i = 0; i < colA.length; i++) {
      if (!colA[i][0]) { newRow = startRowHist + i; break; }
    }
  } else {
    newRow = startRowHist;
  }

  const quantity = parseFloat(data.qty) || 0;
  const totalCostEur = parseFloat(data.costEur) || 0;
  const totalCostUsd = parseFloat(data.costUsd) || 0;
  let unitPrice = quantity !== 0 ? totalCostEur / quantity : 0;

  // Ticker Normalization
  let finalTicker = data.ticker.trim();
  if (finalTicker.toLowerCase() === "cash") {
    finalTicker = "Cash"; 
  } else {
    finalTicker = finalTicker.toUpperCase();
  }

  // --- OPTIMIZED BATCH WRITE TO HISTORY ---
  let histRowData = new Array(13).fill("");
  
  histRowData[0] = new Date(data.date);
  histRowData[1] = finalTicker;
  histRowData[2] = data.action;
  histRowData[3] = quantity;

  if (isCrypto) {
      histRowData[4] = totalCostEur;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId; // M = 13
  } else {
      let assetClass = (data.action === "Withdrawal") ? "x" : ((data.type === "Stocks") ? "Stock" : "ETF");
      histRowData[4] = assetClass;
      histRowData[5] = unitPrice;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId; // M = 13
  }

  // Scrive tutta la riga in un colpo solo
  sheet.getRange(newRow, 1, 1, 13).setValues([histRowData]);

  // --- PART 2: EXPENSES TRACKER (Withdrawals Only) ---
  if (data.action === "Withdrawal") {
    const txDate = new Date(data.date);
    const year = txDate.getFullYear();
    const expSheetName = "Expenses Tracker " + year;
    const expSheet = ss.getSheetByName(expSheetName);
    
    if (expSheet && data.bankCol) {
      const startRowExp = 20;
      
      // Find empty row logic
      const colB = expSheet.getRange(startRowExp, 2, Math.max(1, expSheet.getLastRow() - startRowExp + 2), 1).getValues();
      let targetRow = startRowExp;
      for (let i = 0; i < colB.length; i++) {
        if (!colB[i][0]) { targetRow = startRowExp + i; break; }
      }

      const category = isCrypto ? "Crypto" : "Stocks";
      const userNote = (data.note && data.note.trim() !== "") ? data.note : "";

      // Dynamically find "Controllo Automatismi" column
      const headerRowExp = 2;
      const maxCols = expSheet.getLastColumn();
      const headersExp = expSheet.getRange(headerRowExp, 1, 1, maxCols).getValues()[0];
      const syncColIndexExp = headersExp.indexOf("Controllo Automatismi") + 1; 
      
      // OPTIMIZED BATCH WRITE FOR EXPENSES
      // Create an array large enough to hold all data up to the sync column
      const numCols = Math.max(maxCols, syncColIndexExp);
      let expRowData = new Array(numCols).fill(""); 
      
      expRowData[0] = txDate;         // A: Date
      expRowData[1] = 'Investment';   // B: Type
      expRowData[2] = category;       // C: Category
      expRowData[3] = userNote;       // D: Note
      
      const colIndex = parseInt(data.bankCol);
      if (!isNaN(colIndex) && colIndex > 0) {
        expRowData[colIndex - 1] = Math.abs(totalCostEur); // -1 because Array is 0-indexed
      }
      
      if (syncColIndexExp > 0) {
        expRowData[syncColIndexExp - 1] = transactionId;
      }

      // Write everything in a single API call
      expSheet.getRange(targetRow, 1, 1, numCols).setValues([expRowData]);
    }
  }
  return "Investment Saved (" + data.type + ")";
}

/**
 * BACKEND: Gestisce Vendita o Aggiornamento Stima Asset
 */
function manageAssetBackend(payload) {
  const { id, action, newVal, sellPrice, bank, closeDebt } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssets = ss.getSheetByName("DB_RealAssets");
  const aData = dbAssets.getDataRange().getValues();
  
  let aRow = -1; let asset = null;
  for(let i=1; i<aData.length; i++) {
     if(aData[i][0] === id) { aRow = i+1; asset = { name: aData[i][2], type: aData[i][1], linkedDebt: aData[i][10] }; break; }
  }
  if(!asset) throw new Error("Asset non trovato");

  if(action === 'Update') {
     if(asset.type !== "Real Estate") throw new Error("Puoi aggiornare manualmente solo gli Immobili.");
     dbAssets.getRange(aRow, 6).setValue(parseFloat(newVal)); // Colonna F: Purchase_Price / Market_Value
     return "Valore di mercato aggiornato a €" + newVal;
  }

  if(action === 'Sell') {
     dbAssets.getRange(aRow, 12).setValue("Sold"); // Imposta Status su Sold
     let cashIn = parseFloat(sellPrice);
     let notes = `Vendita Asset: ${asset.name}`;

     // Gestione Mutuo Collegato (Estinzione Contestuale)
     if(asset.linkedDebt && closeDebt) {
        try {
           // Chiama in background l'estinzione totale, ma senza fargli registrare la spesa (lo compensiamo qui)
           // Per semplicità logica, calcoliamo il residuo e lo sottraiamo dall'incasso
           const dbDebts = ss.getSheetByName("DB_Debts");
           let dData = dbDebts.getDataRange().getValues();
           for(let j=1; j<dData.length; j++) {
              if(dData[j][0] === asset.linkedDebt) {
                 let d = { start: new Date(dData[j][2]), prin: parseFloat(dData[j][3]), rate: parseFloat(dData[j][4]), yrs: parseFloat(dData[j][5]) };
                 let mPassed = (new Date().getFullYear() - d.start.getFullYear())*12 + (new Date().getMonth() - d.start.getMonth());
                 let mR = d.rate/12; let tM = d.yrs*12;
                 let out = d.prin;
                 if(mPassed>0 && mPassed<tM) {
                    let p = d.prin * (mR * Math.pow(1+mR, tM)) / (Math.pow(1+mR, tM) - 1);
                    if(isNaN(p)) p = d.prin/tM;
                    out = p * (1 - Math.pow(1+mR, -(tM-mPassed))) / mR;
                 }
                 cashIn -= out; // Sottrae il debito residuo dall'incasso netto
                 notes += ` (Estinto mutuo di €${out.toFixed(2)})`;
                 dbDebts.getRange(j+1, 8).setValue("Paid_Off");
                 break;
              }
           }
        } catch(e) { Logger.log("Errore estinzione mutuo collegato: " + e.message); }
     }

     // Registra l'Entrata netta
     let amountsObj = {}; amountsObj[bank] = cashIn;
     addTransaction({ type: 'Income', category: "Investimenti", details: notes, amounts: amountsObj });
     
     return `Asset venduto. Incasso netto (post-mutuo) accreditato: €${cashIn.toFixed(2)}`;
  }
}