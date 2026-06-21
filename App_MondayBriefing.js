/**
 * App_MondayBriefing.js
 * Sends a weekly strategy email every Monday morning.
 * Uses Universal AI Router for robust fallback (OpenRouter -> Gemini).
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
    to: getNotifyEmail(), // FIX (2.17): email centralizzata
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
 * Generates the briefing using Universal AI Router.
 */
function generateBriefingAI(market, watchlist) {
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
    
    FORMAT: HTML (clean div/p tags), professional but energetic tone. Do not wrap in markdown blocks like \`\`\`html.
  `;

  console.log("Monday Briefing AI: Attempting OPENROUTER...");
  let responseText = fetchUniversalAI(prompt, 'OPENROUTER');

  if (!responseText) {
    console.log("Monday Briefing AI: OPENROUTER failed. Falling back to GEMINI...");
    responseText = fetchUniversalAI(prompt, 'GEMINI');
  }

  if (responseText) {
    // Clean up any accidental markdown blocks that the AI might insert
    return responseText.replace(/```html/gi, "").replace(/```/g, "").trim();
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