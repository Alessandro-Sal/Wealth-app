/**
 * Retrieves Tax Optimization data including estimated tax bill and loss harvesting potential.
 * @return {Object} Structured data for the UI.
 */
function getTaxOptimizationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tax Optimization");
  if (!sheet) return { error: "Sheet 'Tax Optimization' not found" };

  // Smart helper to parse numbers in both US and IT formats safely
  const safeParseFloat = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    
    let s = String(val).replace(/[€\s]/g, "");
    
    // If it contains both dot and comma (e.g., 1.250,50 or 1,250.50)
    if (s.includes(".") && s.includes(",")) {
      if (s.lastIndexOf(".") > s.lastIndexOf(",")) {
        // US Format: 1,250.50 -> remove commas
        s = s.replace(/,/g, "");
      } else {
        // IT Format: 1.250,50 -> remove dots, comma becomes dot
        s = s.replace(/\./g, "").replace(",", ".");
      }
    } else if (s.includes(",")) {
      // Only comma, assume IT decimal: 40,30
      s = s.replace(",", ".");
    }
    // If it contains only a dot (e.g. 40.30), parseFloat already reads it correctly
    
    let num = parseFloat(s);
    return isNaN(num) ? 0 : num;
  };

  // 1. Read data for Tax Bill (from A2:C15 dynamically finding keys)
  const billRange = sheet.getRange("A2:C15").getValues();
  let taxBillData = [];
  for (let i = 0; i < billRange.length; i++) {
    let metric = billRange[i][0];
    let val = billRange[i][1];
    let desc = billRange[i][2];
    
    if (!metric || metric.toString().includes("----")) continue;
    
    // Select only known rows to avoid garbage
    if (["Realized Gains (Compensable)", "Realized Losses (This Year)", "Net Annual Position", "Previous Basket (Zainetto)", "TAXABLE BASE", "ESTIMATED TAX DUE", "Residual Basket (Future)", "New Basket Generated"].includes(metric)) {
        taxBillData.push({
            metric: metric,
            value: safeParseFloat(val),
            desc: desc
        });
    }
  }

  // 2. Read Tax Loss Harvesting Scanner dynamically starting from row 13
  const lastRow = sheet.getLastRow();
  let harvestingData = [];
  let totalPotential = 0;
  
  if (lastRow >= 13) {
    // Start from row 13 safely up to the end of the sheet
    const harvRange = sheet.getRange(13, 1, lastRow - 12, 5).getValues();
    let isHarvestingSection = false;
    
    for (let i = 0; i < harvRange.length; i++) {
      let ticker = harvRange[i][0];
      if (!ticker) continue;
      
      // When we meet the header, we trigger data reading
      if (ticker === "Ticker") {
        isHarvestingSection = true;
        continue;
      }
      
      if (!isHarvestingSection) continue;
      
      // Found the total, record it and break the loop
      if (ticker === "TOTAL POTENTIAL") {
        totalPotential = safeParseFloat(harvRange[i][3]);
        break;
      }
      
      // Exclude empty message and add assets
      if (ticker !== "No losing positions found.") {
        harvestingData.push({
          ticker: ticker,
          qty: harvRange[i][1],
          unrealized: safeParseFloat(harvRange[i][2]),
          minus: safeParseFloat(harvRange[i][3]),
          note: harvRange[i][4]
        });
      }
    }
  }

  return {
    bill: taxBillData,
    harvesting: harvestingData,
    totalPotential: totalPotential
  };
}