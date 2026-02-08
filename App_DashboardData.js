/**
 * Retrieves key metrics from the "Net Worth OGGI" sheet.
 * Maps specific rows for Total and Liquid Net Worth based on user configuration.
 * * @return {Object} Dashboard data object containing summary stats and section-specific details.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Net Worth OGGI");
  if (!sheet) return { error: "Sheet 'Net Worth OGGI' not found" };

  // Helper to read Value (Col B=2) and Percentage (Col C=3) for top tables
  const getRowData = (row) => {
    const val = sheet.getRange(row, 2).getDisplayValue(); 
    const raw = sheet.getRange(row, 2).getValue();
    const pct = sheet.getRange(row, 3).getDisplayValue(); 
    return { amount: val, raw: (typeof raw === 'number' ? raw : 0), percent: pct };
  };

  const getSectionData = (startRow) => {
    return {
      unrealized: { 
        amount: sheet.getRange(startRow, 2).getDisplayValue(),     
        percent: sheet.getRange(startRow, 3).getDisplayValue() 
      },
      realized: { 
        amount: sheet.getRange(startRow + 1, 2).getDisplayValue(), 
        percent: sheet.getRange(startRow + 1, 3).getDisplayValue() 
      },
      balance: { 
        amount: sheet.getRange(startRow + 2, 2).getDisplayValue(), 
        percent: sheet.getRange(startRow + 2, 3).getDisplayValue() 
      },
      invested: {
        amount: sheet.getRange(startRow + 3, 2).getDisplayValue(), 
        percent: "" 
      }
    };
  };

  return {
    // --- MAPPING CORRECTION ---
    
    // Liquid NW at Row 26 (Col A=1 EUR, Col B=2 USD)
    liquidNetWorth: sheet.getRange(26, 1).getDisplayValue(),    
    liquidNetWorthUSD: sheet.getRange(26, 2).getDisplayValue(), 
    
    // Total NW at Row 24 (Col A=1 EUR, Col B=2 USD)
    totalNetWorth: sheet.getRange(24, 1).getDisplayValue(),     
    totalNetWorthUSD: sheet.getRange(24, 2).getDisplayValue(),  

    // ----------------------------

    summary: { 
      etfs: getRowData(2),      
      stocks: getRowData(3),    
      cash: getRowData(4),      
      cashEq: getRowData(5),    
      crypto: getRowData(6),    
      others: getRowData(7)     
    },

    cryptoSection: {
      main: getRowData(6), 
      ...getSectionData(9) 
    },
    stocksSection: {
      main: getRowData(3),
      ...getSectionData(14)
    },
    etfSection: {
      main: getRowData(2),
      ...getSectionData(19)
    }
  };
}