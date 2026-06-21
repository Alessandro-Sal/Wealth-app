/**
 * App_Reporting.js
 * - Generates forensic monthly reports using AI.
 * - Sends the report via Email.
 * - Saves the report as PDF on Google Drive.
 * - Inserts the link to the PDF directly into the "NW analitico" sheet.
 */

const REPORT_CONFIG = {
  emailSubject: "📊 Wealth Report - Forensic Analysis",
  sheetName: "NW analitico",
  reportRowLabel: "Report AI", // Ensure this label exists in Column A (or it will be created)
  driveFolderName: "Wealth Reports Archive"
};

/**
 * MONTHLY TRIGGER (Run this on the 1st of the month or manually)
 */
function sendMonthlyWealthReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_CONFIG.sheetName);
  
  if (!sheet) { console.error("Sheet 'NW analitico' missing"); return; }

  // 1. Historical Data (Month over Month comparison)
  const monthlyData = getLastTwoMonthsComparison(sheet);
  
  // If insufficient data, stop
  if (!monthlyData) {
    console.warn("Unable to generate report: insufficient historical data.");
    return;
  }

  // 2. Macro Context (S&P500, VIX)
  let market = { spx: "N/A", vix: "N/A", cryptoTrend: "N/A" };
  try { market = getMarketInsightsData(true); } catch(e) {}

  // 3. Portfolio Specifics (Top Movers)
  let portfolioHighlights = "Asset details unavailable.";
  try {
    const rawPort = getLivePortfolio(); 
    portfolioHighlights = extractPortfolioMovers(rawPort);
  } catch(e) { console.warn("Error reading portfolio details: " + e); }

  // 4. Transactions History (B/S Stocks & Crypto for the target month)
  const transactionsHistory = getMonthlyTransactionsForReporting(monthlyData.targetDate);

  // 5. Advanced AI Analysis Generation
  const aiAnalysis = generateReportAI_Advanced(monthlyData, market, portfolioHighlights, transactionsHistory);

  // 6. Create HTML Template
  const emailHtml = createEmailTemplate(aiAnalysis, monthlyData);

  // 7. SEND EMAIL
  MailApp.sendEmail({
    to: getNotifyEmail(), // FIX (2.17): email centralizzata (vedi App_NotifyConfig.js)
    subject: `${REPORT_CONFIG.emailSubject} - ${monthlyData.currentMonthName}`,
    htmlBody: emailHtml
  });
  
  // 8. SAVE TO DRIVE (PDF) AND LINK TO SHEET
  saveReportToDriveAndLink(sheet, monthlyData, emailHtml);
  
  console.log("✅ Report Sent, Saved as PDF on Drive, and Linked in Sheet.");
}

/**
 * Scans the History sheets to grab transactions performed during the target month.
 */
function getMonthlyTransactionsForReporting(targetDate) {
  if (!targetDate) return "Nessuna data di riferimento trovata.";
  
  const targetMonth = targetDate.getMonth();
  const targetYear = targetDate.getFullYear();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToScan = ["History B/S Stocks", "History B/S Crypto"];
  
  let historyText = "\n=== STORICO TRANSAZIONI DEL MESE (Depositi, Acquisti, Vendite) ===\n";
  let transactionCount = 0;

  sheetsToScan.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    let sheetHistory = [];
    
    // Loop through rows skipping the header (row 0)
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0]; // Col A
      
      // Basic date check
      if (rowDate && (rowDate instanceof Date || !isNaN(new Date(rowDate).getTime()))) {
         const d = new Date(rowDate);
         
         // Match Month and Year
         if (d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
           const ticker = data[i][1]; // Col B (Ticker / Cash)
           const action = data[i][2]; // Col C (Action)
           const qty = data[i][3];    // Col D (Quantity)
           const total = data[i][7];  // Col H (Total amount)
           
           let formattedDate = Utilities.formatDate(d, Session.getScriptTimeZone() || "GMT", "dd/MM/yyyy");
           let formatAmt = formatMoney(parseFinanceValue(total));
           
           sheetHistory.push(`- [${sheetName}] Data: ${formattedDate} | Ticker: ${ticker} | Azione: ${action} | Qty: ${qty} | Totale: ${formatAmt}`);
           transactionCount++;
         }
      }
    }
    
    if (sheetHistory.length > 0) {
      historyText += sheetHistory.join("\n") + "\n";
    }
  });

  if (transactionCount === 0) {
    return historyText + "Nessuna transazione registrata per questo mese.\n";
  }
  return historyText;
}

/**
 * Saves the HTML content as a PDF file on Drive and links it in the Excel sheet.
 */
function saveReportToDriveAndLink(sheet, data, htmlContent) {
  try {
    const fileName = `Report ${data.currentMonthName}.pdf`; 
    
    const folders = DriveApp.getFoldersByName(REPORT_CONFIG.driveFolderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(REPORT_CONFIG.driveFolderName);
    }

    const htmlBlob = Utilities.newBlob(htmlContent, "text/html", "temp.html");
    const pdfBlob = htmlBlob.getAs(MimeType.PDF).setName(fileName);

    const existingFiles = folder.getFilesByName(fileName);
    let file;
    
    if (existingFiles.hasNext()) {
      const oldFile = existingFiles.next();
      oldFile.setTrashed(true);
    }
    
    file = folder.createFile(pdfBlob);
    const fileUrl = file.getUrl();

    const lastRow = sheet.getLastRow();
    const labels = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
    let targetRow = -1;

    for (let i = 0; i < labels.length; i++) {
      if (labels[i][0].trim() === REPORT_CONFIG.reportRowLabel) {
        targetRow = i + 1; 
        break;
      }
    }

    if (targetRow === -1) {
      targetRow = lastRow + 1;
      sheet.getRange(targetRow, 1).setValue(REPORT_CONFIG.reportRowLabel).setFontWeight("bold");
    }

    const targetCol = data.currentColIndex; 
    if (targetCol) {
      const cell = sheet.getRange(targetRow, targetCol);
      cell.setFormula(`=HYPERLINK("${fileUrl}"; "📄 PDF Report")`);
      cell.setHorizontalAlignment("center");
    }

  } catch (e) {
    console.error("Error saving to Drive: " + e.toString());
  }
}

/**
 * Forensic Prompt relying on the new Universal AI Router.
 */
function generateReportAI_Advanced(data, market, portfolioDetails, transactionsHistory) {
  const prompt = `
    RUOLO: Sei un Analista Finanziario Forense (Hedge Fund Risk Manager).
    Il tuo compito NON è fare un riassunto, ma individuare le CAUSE ESATTE delle variazioni patrimoniali.
    
    DEFINIZIONI CRITICHE:
    1. "Liquid NW" (Liquid Net Worth) = PATRIMONIO NETTO INVESTIBILE. Include: Cash, Azioni, ETF, Crypto. NON è solo la liquidità in banca. È il valore di mercato dei tuoi asset liquidabili.
    2. "Confronto": ${data.prevMonthName} vs ${data.currentMonthName}.

    DATI A DISPOSIZIONE:
    
    A) CONTESTO MACRO ECONOMICO:
    - S&P 500: ${market.spx}% | VIX: ${market.vix} | Crypto Trend: ${market.cryptoTrend}

    B) VARIAZIONI MENSILI (Strutturate per Sezione):
    ATTENZIONE: 
    1. I dati sono raggruppati in categorie (es. [Stock Market], [ETFs]).
    2. La categoria [Market Indices (For Comparison)] contiene SOLO indici di mercato (S&P 500, NASDAQ, Oro, VIX). NON sono asset che l'utente possiede. Usali UNICAMENTE per confrontare la performance del portafoglio (es. "Il mercato ha fatto +5%, ma il portafoglio ha reso il +3%").
    
    ${data.csv}

    C) DETTAGLIO ASSET SPECIFICI (I "Colpevoli"):
    ${portfolioDetails}

    D) STORICO TRANSAZIONI DEL MESE (Operatività):
    Contiene gli acquisti, vendite, e depositi/prelievi (Cash) fatti nel mese analizzato.
    Usa questi dati per spiegare ESATTAMENTE dove è andata la liquidità e giustificare i Cash Flow.
    ${transactionsHistory}

    OBIETTIVO DEL REPORT (In Italiano, Formato HTML pulito <div><p>):
    
    1. 🚨 **L'Headline (Il Verdetto)**: 
       Analizza subito la riga "Liquid NW". È salito o sceso rispetto al mese scorso?
    
    2. 🕵️ **Analisi Forense (Perché è successo?)**:
       Devi collegare i totali ai singoli asset e allo Storico Transazioni.
       - SE "[Stock Market]" è sceso E nel dettaglio vedi "AAPL -30%", scrivilo.
       - Giustifica i movimenti di liquidità o l'aumento dei capitali investiti usando la sezione D (es. "L'aumento della liquidità è dovuto alla vendita di Apple e a un deposito di 1000€").
    
    3. 📊 **Performance vs Mercato**:
       Il portafoglio ha retto meglio dell'S&P500 o peggio? Usa la sezione [Market Indices (For Comparison)] per dare contesto (es. "Il mercato Nasdaq è sceso, trascinando il portafoglio").

    4. 💡 **Cash Flow & Risparmio**:
       Analizza rigorosamente le sezioni [Income] (Entrate), [Expenses] (Uscite) e [Savings] (Risparmi). Il risparmio mensile ha compensato le perdite di mercato o viceversa? Hai speso più di quanto hai guadagnato?
       
    FORMATTAZIONE:
    - Usa grassetto (<b>) per i numeri, le macro-sezioni e i nomi degli asset (es. <b>Apple</b> o <b>Stock Market</b>).
    - Sii diretto, brutale e specifico. Niente giri di parole.
  `;

  // Routes the prompt to the new Universal AI configuration.
  const aiResponse = fetchUniversalAI(prompt, 'OPENROUTER', false);

  if (aiResponse) {
    return aiResponse.replace(/```html/g, "").replace(/```/g, "").trim();
  }
  
  return "<p>Error: AI Analysis failed.</p>";
}

/**
 * Extracts Top Movers from the portfolio.
 */
function extractPortfolioMovers(portfolio) {
  if (!portfolio) return "";

  const fmt = (list) => {
    return list
      .sort((a, b) => {
        let valA = parseFloat(String(a.pct).replace('%','')) || 0;
        let valB = parseFloat(String(b.pct).replace('%','')) || 0;
        return valB - valA; 
      })
      .slice(0, 5) 
      .map(a => `- ${a.t}: ${a.val} (Total Return: ${a.pct})`)
      .join("\n");
  };

  let txt = "\n=== LIVE PORTFOLIO DETAILS (Top Holdings & Performance) ===\n";
  if (portfolio.stocks && portfolio.stocks.length > 0) txt += "**STOCKS:**\n" + fmt(portfolio.stocks) + "\n";
  if (portfolio.crypto && portfolio.crypto.length > 0) txt += "**CRYPTO:**\n" + fmt(portfolio.crypto) + "\n";
  return txt;
}

/**
 * Robust DATE Logic (Past + Current Month).
 * Modified to accurately track Sections and isolate Market Indices.
 */
function getLastTwoMonthsComparison(sheet) {
  const lastCol = sheet.getLastColumn();
  
  const labels = sheet.getRange(1, 1, 20, 1).getDisplayValues(); 
  let nwRowIndex = 2; 
  for (let i = 0; i < labels.length; i++) {
    if (labels[i][0].toLowerCase().includes("net worth") || labels[i][0].toLowerCase().includes("total nw")) {
      nwRowIndex = i + 1;
      break;
    }
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const nwValues = sheet.getRange(nwRowIndex, 1, 1, lastCol).getValues()[0];

  let validCols = [];
  const today = new Date();
  today.setHours(0,0,0,0);
  
  const isFirstDayOfMonth = today.getDate() === 1;

  for (let i = 1; i < lastCol; i++) {
    let headerStr = String(headers[i]).trim();
    let val = nwValues[i];
    let colDate = null;

    if (headerStr.includes("/")) {
      let parts = headerStr.split("/");
      if (parts.length === 3) {
        colDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    } 
    else if (/^\d{4}$/.test(headerStr)) {
      colDate = new Date(parseInt(headerStr), 11, 31);
    }

    if (colDate && !isNaN(colDate.getTime())) {
      let isCurrentMonthCol = colDate.getMonth() === today.getMonth() && colDate.getFullYear() === today.getFullYear();
      let isValidDate = false;

      if (isFirstDayOfMonth) {
        isValidDate = (colDate < today) && !isCurrentMonthCol;
      } else {
        isValidDate = colDate <= today;
      }

      if (isValidDate) {
        let hasData = (val !== "" && val !== null && val !== undefined);
        if (hasData) {
          validCols.push({ index: i + 1, name: headerStr, date: colDate });
        }
      }
    }
  }

  validCols.sort((a, b) => a.date - b.date);

  if (validCols.length < 2) {
    console.warn(`Only found ${validCols.length} valid periods.`);
    return null; 
  }

  const current = validCols[validCols.length - 1]; 
  const prev = validCols[validCols.length - 2];    

  console.log(`Selected Comparison: ${prev.name} vs ${current.name}`);

  const maxRows = Math.max(100, sheet.getLastRow());
  const rangeA = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  const rangePrev = sheet.getRange(1, prev.index, maxRows, 1).getDisplayValues();
  const rangeCurr = sheet.getRange(1, current.index, maxRows, 1).getDisplayValues();

  const mainSections = [
    "Cryptocurrency", "Stock Market", "ETFs", "Stocks", 
    "Cumulated Pension", "Expenses", "Income", "Savings"
  ];
  const marketIndicesMatches = [
    "NYSEARCA:GLD", "VIX", "TNX", "S&P 500", "NASDAQ", "Dow Jones", "Russel 2000"
  ];
  
  let currentSection = "General Net Worth";
  let csvData = "SECTION | METRIC | " + prev.name + " | " + current.name + " | DELTA\n";

  for (let i = 1; i < maxRows; i++) {
    let label = rangeA[i][0].trim();
    if (!label || label === "") continue;
    if (label === "(V)") continue; // Remove useless spacer rows

    // Detect sections or indices
    if (mainSections.includes(label)) {
      currentSection = label;
    } else if (marketIndicesMatches.some(idx => label.includes(idx))) {
      currentSection = "Market Indices (For Comparison)";
    }

    let valPrevStr = rangePrev[i][0].trim();
    let valCurrStr = rangeCurr[i][0].trim();

    let valPrev = parseFinanceValue(valPrevStr);
    let valCurr = parseFinanceValue(valCurrStr);
    
    let diff = valCurr - valPrev;
    let diffStr = "";
    if (diff > 0) diffStr = "+" + formatMoney(diff);
    else if (diff < 0) diffStr = formatMoney(diff);
    else diffStr = "€ 0";

    csvData += `[${currentSection}] | ${label} | ${valPrevStr || '-'} | ${valCurrStr || '-'} | ${diffStr}\n`;
  }

  return { 
    csv: csvData, 
    currentMonthName: current.name, 
    prevMonthName: prev.name,
    currentColIndex: current.index,
    targetDate: current.date // Export Date to fetch B/S history
  };
}

function createEmailTemplate(aiContent, data) {
  return `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; max-width: 650px; margin: 0 auto; background-color: #ffffff; padding: 20px; border: 1px solid #e0e0e0;">
      <div style="border-bottom: 3px solid #0056b3; padding-bottom: 15px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #0056b3; text-transform: uppercase; font-size: 20px;">Analisi Patrimoniale Forense</h2>
        <p style="margin: 5px 0 0; color: #666; font-size: 14px;">Periodo: ${data.prevMonthName} ➡ ${data.currentMonthName}</p>
      </div>
      <div style="font-size: 15px; line-height: 1.6; color: #2c3e50;">
        ${aiContent}
      </div>
      <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #eee; font-size: 11px; color: #999; text-align: right;">
        Generato automaticamente da Wealth Manager AI
      </div>
    </div>
  `;
}

function parseFinanceValue(str) {
  if (!str) return 0;
  let s = String(str).replace(/[€$£%\s]/g, '');
  if (s.includes('.') && s.includes(',')) { s = s.replace(/\./g, '').replace(',', '.'); } 
  else if (s.includes(',')) { s = s.replace(',', '.'); }
  return parseFloat(s) || 0;
}

function formatMoney(num) {
  return "€ " + num.toLocaleString("it-IT", {minimumFractionDigits: 0, maximumFractionDigits: 0});
}