/**
 * App_MondayBriefing.js
 * Sends a weekly strategy email every Monday morning.
 * USES ROBUST AI FALLBACK to prevent timeouts.
 */

function sendMondayBriefing() {
  // 1. Fetch Macro Data (S&P500, VIX, etc.)
  let market = { spx: "N/A", vix: "N/A", cryptoTrend: "N/A" };
  try { market = getMarketInsightsData(true); } catch(e) { console.warn("Market data error: " + e); }

  // 2. Fetch Watchlist Data (Top picks from the sheet)
  const watchlist = getWatchlistTopPicks(); 

  // 3. Generate AI Content (Robust Version)
  const aiContent = generateBriefingAI(market, watchlist);

  // 4. Send Email
  MailApp.sendEmail({
    to: "alessandro.saladino01@gmail.com", // ⚠️ Your Email
    subject: "☕ Monday Market Briefing - Weekly Strategy",
    htmlBody: createBriefingTemplate(aiContent)
  });
  
  console.log("Monday Briefing sent successfully.");
}

/**
 * Helper: Reads the Watchlist sheet.
 * Assumes data starts from Row 2.
 */
function getWatchlistTopPicks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Watchlist");
  if (!sheet) return "No Watchlist sheet found.";
  
  // Reads range: Row 2, Col A to F (adjust based on your real columns)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Watchlist empty.";

  const data = sheet.getRange(2, 1, Math.min(10, lastRow - 1), 6).getDisplayValues();
  
  return data
    .filter(r => r[0] !== "") // Filter out empty rows
    .map(r => `- ${r[0]} (${r[1]}): Target Price ${r[4]}`)
    .join("\n");
}

/**
 * Generates the briefing using Gemini with Model Fallback.
 */
function generateBriefingAI(market, watchlist) {
  const API_KEY = GEMINI_API_KEY; 
  
  // ROBUST MODEL SEQUENCE
  const MODELS = [
    "gemini-2.0-flash",      // Priority 1
    "gemini-flash-latest",   // Priority 2
    "gemini-2.0-flash-lite"  // Priority 3 (Backup)
  ];

  const prompt = `
    ROLE: You are a Senior Investment Strategist. It is Monday Morning.
    
    MACRO CONTEXT:
    - S&P500: ${market.spx}% | VIX: ${market.vix} | Crypto Trend: ${market.cryptoTrend}
    
    USER WATCHLIST:
    ${watchlist}

    OBJECTIVE:
    Write a short, motivating email to start the trading week (In Italian).
    1. **Market Sentiment**: Is it Fear or Greed based on VIX?
    2. **Weekly Focus**: Be cautious or aggressive?
    3. **Stock Pick**: Select 1 stock from the watchlist above and explain why to watch it.
    
    FORMAT: HTML (clean div/p tags), professional but energetic tone.
  `;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify(payload), 
    muteHttpExceptions: true 
  };

  // --- RETRY LOOP ---
  for (let i = 0; i < MODELS.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[i]}:generateContent?key=${API_KEY}`;
      console.log(`Attempting Briefing with model: ${MODELS[i]}`);
      
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        return text.replace(/```html/g, "").replace(/```/g, "").trim();
      }
    } catch(e) {
      console.warn(`Model ${MODELS[i]} failed. Trying next.`);
    }
  }

  return "<p>Error: AI models are currently unavailable. Please check API Key or Quota.</p>";
}

function createBriefingTemplate(content) {
  return `<div style="font-family: Arial, sans-serif; color: #333; padding: 20px; border-left: 5px solid #2ecc71; background-color: #f9f9f9;">
    <h2 style="color: #27ae60; margin-top: 0;">☕ Il Tuo Briefing del Lunedì</h2>
    ${content}
    <br><hr><small style="color: #999;">Generato da Wealth Manager AI</small>
  </div>`;
}