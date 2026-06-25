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
  // FIX (1.10): LockService come per addTransaction (evita righe sovrascritte da chiamate concorrenti).
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return "Error: Sistema occupato, riprova tra qualche secondo.";
  }

  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let warnings = []; // FIX (3.1): raccoglie i fallimenti parziali da segnalare all'utente

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
        let rateDec = parseFloat(data.rate) / 100;
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
          principal, rateDec, years, Number(monthlyPayment.toFixed(2)), "Active", ""
        ]);

        try {
            addFixedExpenseToSheet({
              cat: "Housing",
              note: "Rata " + data.name,
              amt: Number(monthlyPayment.toFixed(2)),
              startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd"),
              endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd"),
              payDay: start.getDate(),
              bankCol: data.bankCol, 
              isSplit: false
            });
        } catch(e) {
            // FIX (3.1): non ingoiare l'errore in silenzio
            Logger.log("Errore creazione spesa fissa mutuo: " + e.message);
            warnings.push("rata mensile del mutuo non creata");
        }
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
      
      rowData[5] = marketVal; 
      rowData[10] = linkedDebtId; 

      // --- Sottrae l'anticipo dal conto in IN NEGATIVO ---
      let upfrontCash = marketVal - loanAmt;
      if (upfrontCash > 0 && data.bankCol) {
          let amountsObj = {};
          amountsObj[data.bankCol] = -Math.abs(upfrontCash); // FIX: Importo Negativo
          try {
              addTransaction({
                  type: 'Investment', 
                  category: 'Real Estate',
                  details: 'Acquisto / Anticipo: ' + data.name,
                  amounts: amountsObj
              });
          } catch(e) { Logger.log("Errore cassa: " + e.message); warnings.push("addebito dell'anticipo sul conto non riuscito"); }
      }

    } else if (data.type === "Bond") {
      let nominal = parseFloat(data.nominal);
      let price = parseFloat(data.price);
      let totalCost = (nominal * price) / 100;
      
      rowData[5] = totalCost; 
      rowData[6] = nominal; 
      rowData[7] = parseFloat(data.coupon) / 100;
      rowData[9] = data.tax === "true" ? "WhiteList" : "Standard"; 
      rowData[12] = data.isin; 

      // --- Sottrae il costo del Bond IN NEGATIVO ---
      if (totalCost > 0 && data.bankCol) {
          let amountsObj = {};
          amountsObj[data.bankCol] = -Math.abs(totalCost); // FIX: Importo Negativo
          try {
              addTransaction({
                  type: 'Investment', 
                  category: 'Bond',
                  details: 'Acquisto Bond: ' + data.name,
                  amounts: amountsObj
              });
          } catch(e) { Logger.log("Errore cassa bond: " + e.message); warnings.push("addebito del bond sul conto non riuscito"); }
      }
    }
    
    dbAssets.appendRow(rowData);
    // FIX (3.1): se qualche scrittura collaterale e' fallita, lo si segnala invece di
    // mostrare un "Added!" rassicurante ma fuorviante.
    const baseMsg = data.type === "RealEstate" ? "Real Estate & Mortgage Added!" : "Bond Added!";
    return warnings.length
      ? (baseMsg + " ⚠️ Attenzione: " + warnings.join("; ") + " — verifica il saldo del conto.")
      : baseMsg;
  }

  const transactionId = "ID_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
  const isCrypto = data.type === "Crypto";
  const sheetName = isCrypto ? "History B/S Crypto" : "History B/S Stocks";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet missing (" + sheetName + ")";

  const lastRow = sheet.getLastRow();
  let newRow = lastRow + 1;
  const startRowHist = 3;
  
  if (lastRow >= startRowHist) {
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

  let finalTicker = data.ticker.trim();
  if (finalTicker.toLowerCase() === "cash") finalTicker = "Cash"; 
  else finalTicker = finalTicker.toUpperCase();

  let histRowData = new Array(13).fill("");
  histRowData[0] = new Date(data.date);
  histRowData[1] = finalTicker;
  histRowData[2] = data.action;
  histRowData[3] = quantity;

  if (isCrypto) {
      histRowData[4] = totalCostEur;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId;
  } else {
      let assetClass = (data.action === "Withdrawal") ? "x" : ((data.type === "Stocks") ? "Stock" : "ETF");
      histRowData[4] = assetClass;
      histRowData[5] = unitPrice;
      histRowData[6] = totalCostUsd;
      histRowData[12] = transactionId;
  }

  if (isCrypto) {
      sheet.getRange(newRow, 1, 1, 7).setValues([histRowData.slice(0, 7)]);
  } else {
      sheet.getRange(newRow, 1, 1, 7).setValues([histRowData.slice(0, 7)]);
  }
  sheet.getRange(newRow, 13, 1, 1).setValue(transactionId);

  if (data.action === "Withdrawal") {
    const txDate = new Date(data.date);
    const expSheetName = "Expenses Tracker " + txDate.getFullYear();
    const expSheet = ss.getSheetByName(expSheetName);
    
    if (expSheet && data.bankCol) {
      const startRowExp = 20;
      const colB = expSheet.getRange(startRowExp, 2, Math.max(1, expSheet.getLastRow() - startRowExp + 2), 1).getValues();
      let targetRow = startRowExp;
      for (let i = 0; i < colB.length; i++) {
        if (!colB[i][0]) { targetRow = startRowExp + i; break; }
      }

      const category = isCrypto ? "Crypto" : "Stocks";
      const userNote = (data.note && data.note.trim() !== "") ? data.note : "";
      const headersExp = expSheet.getRange(2, 1, 1, expSheet.getLastColumn()).getValues()[0];
      const syncColIndexExp = headersExp.indexOf("Controllo Automatismi") + 1; 
      
      const numCols = Math.max(expSheet.getLastColumn(), syncColIndexExp);
      let expRowData = new Array(numCols).fill(""); 
      
      expRowData[0] = txDate;         
      expRowData[1] = 'Investment';   
      expRowData[2] = category;       
      expRowData[3] = userNote;       
      
      const colIndex = parseInt(data.bankCol);
      if (!isNaN(colIndex) && colIndex > 0) expRowData[colIndex - 1] = Math.abs(totalCostEur); 
      if (syncColIndexExp > 0) expRowData[syncColIndexExp - 1] = transactionId;

      expSheet.getRange(targetRow, 1, 1, numCols).setValues([expRowData]);
    }
  }
  return "Investment Saved (" + data.type + ")";
  } finally {
    lock.releaseLock(); // FIX (1.10): rilascia sempre il lock, su ogni percorso di uscita
  }
}

/**
 * BACKEND: Gestisce Vendita o Aggiornamento Stima Asset
 */
function manageAssetBackend(payload) {
  const { id, action, newVal, sellPrice, bank, closeDebt } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbAssets = ss.getSheetByName("DB_RealAssets");
  const aData = dbAssets.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let aRow = -1; let asset = null;
  for(let i=1; i<aData.length; i++) {
     if(aData[i][0] === id) { aRow = i+1; asset = { name: aData[i][2], type: aData[i][1], linkedDebt: aData[i][10] }; break; }
  }
  if(!asset) throw new Error("Asset non trovato");

  if(action === 'Update') {
     if(asset.type !== "Real Estate") throw new Error("Puoi aggiornare manualmente solo gli Immobili.");
     dbAssets.getRange(aRow, 6).setValue(parseFloat(newVal)); 
     return "Valore di mercato aggiornato a €" + newVal;
  }

  if(action === 'Sell') {
     dbAssets.getRange(aRow, 12).setValue("Sold"); 
     let cashIn = parseFloat(sellPrice);
     let notes = `Vendita Asset: ${asset.name}`;

     if(asset.linkedDebt && closeDebt) {
        try {
           const dbDebts = ss.getSheetByName("DB_Debts");
           const dbFixed = ss.getSheetByName("Config_FixedExpenses"); 
           let dData = dbDebts.getDataRange().getValues();
           
           for(let j=1; j<dData.length; j++) {
              if(dData[j][0] === asset.linkedDebt) {
                 let d = { 
                     start: new Date(dData[j][2]), prin: parseFloat(dData[j][3]), 
                     rate: parseFloat(dData[j][4]), yrs: parseFloat(dData[j][5]), balloon: parseFloat(dData[j][8]) || 0 
                 };
                 let mPassed = (new Date().getFullYear() - d.start.getFullYear())*12 + (new Date().getMonth() - d.start.getMonth());
                 let mR = d.rate/12; let tM = d.yrs*12;
                 let out = d.prin;
                 
                 if(mPassed>0 && mPassed<tM) {
                    let p = 0;
                    if (mR === 0) p = (d.prin - d.balloon) / tM;
                    else p = ((d.prin * Math.pow(1+mR, tM)) - d.balloon) * mR / (Math.pow(1+mR, tM) - 1);
                    
                    let remM = tM - mPassed;
                    if (mR === 0) out = (p * remM) + d.balloon;
                    else out = (p * (1 - Math.pow(1+mR, -remM)) / mR) + (d.balloon * Math.pow(1+mR, -remM));
                 }
                 
                 cashIn -= out; 
                 notes += ` (Estinto mutuo di €${out.toFixed(2)})`;
                 dbDebts.getRange(j+1, 8).setValue("Paid_Off");
                 break;
              }
           }
           
           // FIX BUG CRITICO: Usa asset.name per spegnere la spesa fissa (non il nome generico del debito!)
           if(dbFixed && asset.name) {
              let fData = dbFixed.getDataRange().getValues();
              for(let i=1; i<fData.length; i++) {
                 // Cerca la nota che contiene il nome della casa es. "Rata Casa Nuova"
                 if(String(fData[i][2]).includes(asset.name) && fData[i][8] !== true) {
                     dbFixed.getRange(i+1, 8).setValue(todayStr);
                 }
              }
           }
        } catch(e) { Logger.log("Errore estinzione mutuo collegato: " + e.message); }
     }

     let amountsObj = {}; amountsObj[bank] = Math.abs(cashIn);
     addTransaction({ 
         type: 'Income', 
         category: "Other Income", // <--- FIX CATEGORIA
         details: notes, 
         amounts: amountsObj 
     });
     
     return `Asset venduto. Incasso netto accreditato: €${cashIn.toFixed(2)}`;
  }
}