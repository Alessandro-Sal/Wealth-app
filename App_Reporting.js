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

  // 4. Advanced AI Analysis Generation
  const aiAnalysis = generateReportAI_Advanced(monthlyData, market, portfolioHighlights);

  // 5. Create HTML Template
  const emailHtml = createEmailTemplate(aiAnalysis, monthlyData);

  // 6. SEND EMAIL
  MailApp.sendEmail({
    to: "alessandro.saladino01@gmail.com", // ⚠️ Insert your email here
    subject: `${REPORT_CONFIG.emailSubject} - ${monthlyData.currentMonthName}`,
    htmlBody: emailHtml
  });
  
  // 7. SAVE TO DRIVE (PDF) AND LINK TO SHEET
  saveReportToDriveAndLink(sheet, monthlyData, emailHtml);
  
  console.log("✅ Report Sent, Saved as PDF on Drive, and Linked in Sheet.");
}

/**
 * Saves the HTML content as a PDF file on Drive and links it in the Excel sheet.
 */
function saveReportToDriveAndLink(sheet, data, htmlContent) {
  try {
    const fileName = `Report ${data.currentMonthName}.pdf`; // Changed extension to .pdf
    
    // A. Handle Drive Folder
    const folders = DriveApp.getFoldersByName(REPORT_CONFIG.driveFolderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(REPORT_CONFIG.driveFolderName);
    }

    // B. Create/Overwrite File as PDF
    // 1. Convert HTML string to a Blob
    const htmlBlob = Utilities.newBlob(htmlContent, "text/html", "temp.html");
    // 2. Convert Blob to PDF
    const pdfBlob = htmlBlob.getAs(MimeType.PDF).setName(fileName);

    // Check if file exists to update or create new
    const existingFiles = folder.getFilesByName(fileName);
    let file;
    
    if (existingFiles.hasNext()) {
      // Drive doesn't allow direct content overwrite for PDFs easily, 
      // so we trash the old one and create a new one to ensure latest version.
      const oldFile = existingFiles.next();
      oldFile.setTrashed(true);
    }
    
    file = folder.createFile(pdfBlob);
    const fileUrl = file.getUrl();

    // C. Find the "Report AI" row in the sheet
    const lastRow = sheet.getLastRow();
    const labels = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
    let targetRow = -1;

    for (let i = 0; i < labels.length; i++) {
      if (labels[i][0].trim() === REPORT_CONFIG.reportRowLabel) {
        targetRow = i + 1; 
        break;
      }
    }

    // If row not found, create it at the end
    if (targetRow === -1) {
      targetRow = lastRow + 1;
      sheet.getRange(targetRow, 1).setValue(REPORT_CONFIG.reportRowLabel).setFontWeight("bold");
    }

    // D. Paste the link in the correct column
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
 * Forensic Prompt with Increased Memory (8k Tokens)
 * Note: The prompt text remains in Italian to ensure the output report is in Italian.
 */
function generateReportAI_Advanced(data, market, portfolioDetails) {
  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-flash-latest", "gemini-2.0-flash-lite"];

  const prompt = `
    RUOLO: Sei un Analista Finanziario Forense (Hedge Fund Risk Manager).
    Il tuo compito NON è fare un riassunto, ma individuare le CAUSE ESATTE delle variazioni patrimoniali.
    
    DEFINIZIONI CRITICHE:
    1. "Liquid NW" (Liquid Net Worth) = PATRIMONIO NETTO INVESTIBILE. Include: Cash, Azioni, ETF, Crypto. NON è solo la liquidità in banca. È il valore di mercato dei tuoi asset liquidabili.
    2. "Confronto": ${data.prevMonthName} vs ${data.currentMonthName}.

    DATI A DISPOSIZIONE:
    
    A) CONTESTO MACRO ECONOMICO:
    - S&P 500: ${market.spx}% | VIX: ${market.vix} | Crypto Trend: ${market.cryptoTrend}
    (Usa questo per dire se il portafoglio ha sovraperformato o sottoperformato il mercato).

    B) VARIAZIONI MENSILI (Totali):
    ${data.csv}

    C) DETTAGLIO ASSET SPECIFICI (I "Colpevoli"):
    ${portfolioDetails}

    OBIETTIVO DEL REPORT (In Italiano, Formato HTML pulito <div><p>):
    
    1. 🚨 **L'Headline (Il Verdetto)**: 
       Analizza subito la riga "Liquid NW". È salito o sceso?
       Esempio: "Il Patrimonio Investibile è sceso del 3% (€ -4.500) questo mese."
    
    2. 🕵️ **Analisi Forense (Perché è successo?)**:
       Devi collegare i totali ai singoli asset.
       - SE "Stock Market" è sceso E nel dettaglio vedi "AAPL -30%", DEVI SCRIVERE: "Il calo dell'azionario è trainato principalmente dal crollo di Apple (-30%)..."
       - SE "Cryptocurrency" è sceso E vedi "BTC -10%", scrivilo.
       - Distingui tra calo di mercato (Performance) e flussi di cassa (Ho speso troppo).
    
    3. 📊 **Performance vs Mercato**:
       Il portafoglio ha retto meglio dell'S&P500 o peggio? Perché? (Es. "Troppa esposizione Crypto ha aumentato la volatilità").

    4. 💡 **Cash Flow & Risparmio**:
       Analizza "Income", "Expenses" e "Savings". Il risparmio mensile ha compensato le perdite di mercato?
       
    FORMATTAZIONE:
    - Usa grassetto (<b>) per i numeri e i nomi degli asset (es. <b>Apple</b>).
    - Sii diretto, brutale e specifico. Niente giri di parole.
  `;

  const payload = { 
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      maxOutputTokens: 8192, // Max tokens to avoid truncated text
      temperature: 0.7
    }
  };
  
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  for (let i = 0; i < MODELS.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[i]}:generateContent?key=${API_KEY}`;
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        return text.replace(/```html/g, "").replace(/```/g, "").trim();
      }
    } catch (e) { console.warn(`Model ${MODELS[i]} failed.`); }
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

  const isCurrentMonth = (d) => {
    return d.getMonth() === today.getMonth() && d.getFullYear() === today.getFullYear();
  };

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
      // ACCEPT IF: Past Date OR Current Month (even if future day e.g. 28/02 vs 16/02)
      if (colDate <= today || isCurrentMonth(colDate)) {
        let hasData = (val !== "" && val !== null && val !== undefined);
        if (hasData) {
          validCols.push({ index: i + 1, name: headerStr, date: colDate });
        }
      }
    }
  }

  validCols.sort((a, b) => a.date - b.date);

  if (validCols.length < 2) {
    console.warn(`Only found ${validCols.length} valid periods. Need at least Prev and Curr Month.`);
    return null; 
  }

  const current = validCols[validCols.length - 1]; 
  const prev = validCols[validCols.length - 2];    

  console.log(`Selected Comparison: ${prev.name} vs ${current.name}`);

  const lastRow = sheet.getLastRow();
  const rangeA = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  const rangePrev = sheet.getRange(1, prev.index, lastRow, 1).getDisplayValues();
  const rangeCurr = sheet.getRange(1, current.index, lastRow, 1).getDisplayValues();

  let csvData = "METRIC | " + prev.name + " | " + current.name + " | DELTA\n";

  for (let i = 0; i < lastRow; i++) {
    let label = rangeA[i][0];
    let valPrevStr = rangePrev[i][0];
    let valCurrStr = rangeCurr[i][0];

    if (!label || label === "") continue;

    let valPrev = parseFinanceValue(valPrevStr);
    let valCurr = parseFinanceValue(valCurrStr);
    
    if (valPrev !== 0 || valCurr !== 0) {
       let diff = valCurr - valPrev;
       let diffStr = diff > 0 ? "+" + formatMoney(diff) : formatMoney(diff);
       csvData += `${label} | ${valPrevStr} | ${valCurrStr} | ${diffStr}\n`;
    }
  }

  return { 
    csv: csvData, 
    currentMonthName: current.name, 
    prevMonthName: prev.name,
    currentColIndex: current.index // Essential index for linking!
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