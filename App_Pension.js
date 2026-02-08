/**
 * --- PENSION DATA MODULE ---
 * Retrieves and manages pension fund data.
 * Handles a specific sheet structure where Year headers and Data rows are interleaved.
 */

/**
 * Scans the "Pension" sheet to build a history of annual and cumulative contributions.
 * Iterates through rows in steps of 2 (Header Row -> Data Row).
 * Parses Italian number formats (e.g., "1.667,36" -> 1667.36) automatically.
 * * @return {Array<Object>} Array of objects containing { year, annual, cumulative }.
 */
function getPensionData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // WARNING: Ensure the tab name below matches exactly what is in your Sheet
  const sheet = ss.getSheetByName("Pension"); 
  
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Read everything up to column O (15 columns)
  // getDisplayValues is crucial here to capture the formatted text (e.g., "€ 100,00")
  const data = sheet.getRange(1, 1, lastRow, 15).getDisplayValues(); 
  
  let history = [];

  // Iterate with a step of 2 (Year Row -> Data Row)
  // i=0 (2023 Header), i=1 (2023 Data), i=2 (2024 Header)...
  for (let i = 0; i < data.length; i += 2) {
    let rowHeader = data[i];     // Dark Blue Row (Years)
    let rowData = data[i + 1];   // Orange Row (Quotas)

    if (!rowData) break;

    // Column A of the Header row contains the Year (e.g., "2023")
    let year = String(rowHeader[0]).trim(); 
    
    // Internal helper to parse Italian number formats (e.g., "€ 1.667,36" -> 1667.36)
    const parseIta = (val) => {
      if (!val) return 0;
      let s = String(val);
      // specific check: if text contains "Aspettativa" or "x", treat as 0
      if (s.toLowerCase().includes('aspettativa') || s.toLowerCase() === 'x') return 0;
      
      // Remove everything that is NOT a number, minus sign, or comma (removes currency symbols, thousands dots)
      s = s.replace(/[^0-9,-]/g, ''); 
      // Replace decimal comma with dot for JavaScript parsing
      s = s.replace(',', '.');
      
      let num = parseFloat(s);
      return isNaN(num) ? 0 : num;
    };

    // Column N = Index 13 (Annual Total)
    let annualTotal = parseIta(rowData[13]); 
    // Column O = Index 14 (Cumulative Total)
    let cumulativeTotal = parseIta(rowData[14]); 

    // Add to history only if a valid year is found
    if (year && !isNaN(parseFloat(year))) {
       history.push({
        year: year,
        annual: annualTotal,
        cumulative: cumulativeTotal
      });
    }
  }

  return history;
}

/**
 * Adds a single pension transaction to the specific cell for a given Month/Year.
 * Finds the correct row by searching for the Year in Column A, then targeting the row below it.
 * * @param {Object} data - Transaction data.
 * @param {number|string} data.year - The target year (e.g., 2026).
 * @param {number} data.month - The month index (0 = Jan, 11 = Dec).
 * @param {number} data.amount - The amount to add.
 * @return {string} Status message.
 */
function addPensionTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pension");
  if (!sheet) return "Error: Sheet 'Pension' missing";

  // Read the whole sheet to find the Year Row
  const values = sheet.getDataRange().getValues();
  let yearRow = -1;

  // Search for the row containing the Year in Column A
  for (let i = 0; i < values.length; i++) {
    // Compare as strings for safety
    if (String(values[i][0]).trim() === String(data.year)) {
      yearRow = i; // 0-based index
      break;
    }
  }

  if (yearRow === -1) return "Year " + data.year + " not found in Pension sheet";

  // The target row is the one *immediately following* the year row.
  // Apps Script rows are 1-based.
  // yearRow is 0-based index. Year Row = yearRow + 1. Quota Row = yearRow + 2.
  const targetRow = yearRow + 2;

  // Column Calculation: January is Column B (2), December is Column M (13).
  // data.month comes from JS (0 = Jan, 11 = Dec).
  // Logic: 0 -> Col 2 => +2 offset
  const targetCol = parseInt(data.month) + 2;

  // Write value
  sheet.getRange(targetRow, targetCol).setValue(parseFloat(data.amount));

  return "Quota Pension Salvata!";
}