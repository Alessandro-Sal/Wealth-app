/**
 * Retrieves current bank balance data from the "Expenses Tracker" sheet.
 * Implements a dynamic lookup for the current year (e.g., "Expenses Tracker 2026").
 * Includes a fallback mechanism to find *any* available tracker sheet if the current year is missing.
 * * @return {Object|null} An object containing data from columns 5-10, or null if no sheet is found.
 */
function getBankBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Dynamically determine the current year
  const year = new Date().getFullYear(); 
  const sheet = ss.getSheetByName("Expenses Tracker " + year);
  
  // Fallback: If the specific year sheet is missing, try to find the first available "Expenses Tracker..." sheet
  if (!sheet) {
     const sheets = ss.getSheets();
     const expSheet = sheets.find(s => s.getName().startsWith("Expenses Tracker"));
     
     if(!expSheet) return null;
     
     // Retrieve display values from Row 3, Columns 5 to 10
     const data = expSheet.getRange(3, 5, 1, 6).getDisplayValues()[0];
     return { col5: data[0], col6: data[1], col7: data[2], col8: data[3], col9: data[4], col10: data[5] };
  }

  // Retrieve data from the found current year sheet
  const data = sheet.getRange(3, 5, 1, 6).getDisplayValues()[0];
  return { col5: data[0], col6: data[1], col7: data[2], col8: data[3], col9: data[4], col10: data[5] };
}