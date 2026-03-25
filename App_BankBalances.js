/**
 * Retrieves current bank balance data from the "Expenses Tracker" sheet.
 * Implements a dynamic lookup for the current year (e.g., "Expenses Tracker 2026").
 * Includes a fallback mechanism to find *any* available tracker sheet if the current year is missing.
 * @return {Object|null} An object containing data from columns 5-10, plus a structured 'banks' array.
 */
function getBankBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Dynamically determine the current year
  const year = new Date().getFullYear(); 
  let sheet = ss.getSheetByName("Expenses Tracker " + year);
  
  // Fallback: If the specific year sheet is missing, try to find the first available "Expenses Tracker..." sheet
  if (!sheet) {
     const sheets = ss.getSheets();
     sheet = sheets.find(s => s.getName().startsWith("Expenses Tracker"));
     if(!sheet) return null;
  }

  // Legge la Riga 2 (Nomi dei conti) e la Riga 3 (Saldi) per le colonne E-J (5-10)
  const names = sheet.getRange(2, 5, 1, 6).getValues()[0];
  const balances = sheet.getRange(3, 5, 1, 6).getDisplayValues()[0];

  let banksArray = [];
  for (let i = 0; i < 6; i++) {
     let bankName = names[i] ? String(names[i]).trim() : "";
     
     // Aggiunge la banca solo se ha un nome nella riga di intestazione
     if (bankName !== "") {
         banksArray.push({
             name: bankName,
             balance: balances[i] || "0,00"
         });
     }
  }

  return { 
      // Manteniamo la retrocompatibilità assoluta per altre funzioni dell'app
      col5: balances[0], 
      col6: balances[1], 
      col7: balances[2], 
      col8: balances[3], 
      col9: balances[4], 
      col10: balances[5],
      // Array strutturato per la generazione dinamica sulla Dashboard
      banks: banksArray 
  };
}