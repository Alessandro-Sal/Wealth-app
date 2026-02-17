/**
 * App_SniperAlert.js
 * Scans LIVE portfolio for rapid price movements (> 4% or < -4%).
 * - Stocks/ETFs: Uses the "pct" (Percentage Change) column from the sheet.
 * - Crypto: Calculates volatility based on the difference vs. the LAST SCRIPT RUN.
 * Uses ROBUST AI for market commentary.
 */

function runMarketSniper() {
  // 1. Get Live Portfolio Data
  let portfolio;
  try {
     portfolio = getLivePortfolio(); 
  } catch(e) {
     console.error("Sniper Error: Could not fetch portfolio.");
     return;
  }

  // Combine Stocks, ETFs, and Crypto into one array
  const allAssets = [...portfolio.stocks, ...portfolio.etfs, ...portfolio.crypto];
  
  // Retrieve stored prices from the last run (for assets without daily change data, like Crypto)
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastPrices = scriptProperties.getProperties();
  
  let alerts = [];
  let updates = {}; // Object to store current prices for the next run

  // 2. Scan for Volatility
  allAssets.forEach(asset => {
    let changeVal = 0.0;
    let changeStr = "0%";
    
    // --- CASE A: STOCKS & ETFS (Use 'pct' - Percentage Change) ---
    // We check if 'pct' exists and is not a placeholder. We explicitly exclude Crypto here just in case.
    if (asset.pct && asset.pct !== "" && asset.pct !== "0%" && asset.sector !== "Crypto") {
       changeStr = asset.pct;
       // Clean string: "-6,00%" -> -6.00 (replace comma with dot for JS math)
       changeVal = parseFloat(String(changeStr).replace('%', '').replace(',', '.'));
    } 
    // --- CASE B: ASSETS WITHOUT % (e.g., Crypto or missing data) ---
    // We calculate volatility manually by comparing Current Price vs. Last Run Price
    else {
       const cleanPrice = parsePrice(asset.price);
       // Create a unique cache key (e.g., LAST_PRICE_BTC)
       const cacheKey = "LAST_PRICE_" + asset.t.replace(/\s/g, ''); 
       const lastPrice = parseFloat(lastPrices[cacheKey]);

       // Only calculate if we have a valid previous price
       if (lastPrice && lastPrice > 0) {
         changeVal = ((cleanPrice - lastPrice) / lastPrice) * 100;
         changeStr = changeVal.toFixed(2).replace('.', ',') + "% (vs Last Run)";
       }

       // Store the current price to be used as "lastPrice" in the next run
       updates[cacheKey] = String(cleanPrice);
    }

    // THRESHOLDS: Drop < -4% or Pump > +5%
    // Ensure 'changeVal' is a valid number before comparing
    if (!isNaN(changeVal)) {
        if (changeVal <= -4.0) {
          alerts.push(`🔻 <b>${asset.t}</b> is crashing: <span style="color:red">${changeStr}</span>`);
        } else if (changeVal >= 5.0) {
          alerts.push(`🚀 <b>${asset.t}</b> is pumping: <span style="color:green">+${changeStr}</span>`);
        }
    }
  });

  // Save the updated Crypto prices to script properties
  if (Object.keys(updates).length > 0) {
    scriptProperties.setProperties(updates);
  }

  // 3. Send Alert if Triggered
  if (alerts.length > 0) {
    const aiComment = generateSniperAI(alerts);
    
    MailApp.sendEmail({
      to: "alessandro.saladino01@gmail.com",
      subject: `🔥 MARKET MOVER: ${alerts.length} Asset in Motion`,
      htmlBody: `<div style="font-family: Arial; padding: 20px; border: 2px solid #e74c3c; border-radius: 8px;">
        <h2 style="color: #c0392b; margin-top: 0;">🎯 Sniper Alert</h2>
        <ul style="font-size: 16px;">${alerts.map(a => `<li style="margin-bottom: 5px;">${a}</li>`).join('')}</ul>
        <hr style="border: 0; border-top: 1px solid #ccc;">
        <p><i>AI Analyst Opinion:</i><br>${aiComment}</p>
        <br>
        <small style="color:gray;">*Crypto change is calculated based on movement since the last script run.</small>
      </div>`
    });
    console.log(`Sniper Alert Sent: ${alerts.length} assets.`);
  } else {
    console.log("Sniper Scan: Markets are calm.");
  }
}

/**
 * Helper function to parse currency strings like "€ 1.234,56" or "$ 50.40" into a float.
 * Handles both comma and dot decimals based on common European/US formats.
 */
function parsePrice(priceStr) {
  if (!priceStr) return 0;
  // Remove currency symbols, spaces, and keep only digits, commas, dots, and minus signs
  let clean = String(priceStr).replace(/[^\d,\.-]/g, ''); 
  
  // Logic to handle "1.234,56" (IT) vs "1,234.56" (US)
  // If both exist, assume the last separator is the decimal one.
  // Here we assume Italian format (1.000,00) -> remove dots, replace comma with dot.
  if (clean.includes(',') && clean.includes('.')) {
     clean = clean.replace(/\./g, '').replace(',', '.');
  } else if (clean.includes(',')) {
     // If only comma exists (e.g. 12,50), replace with dot
     clean = clean.replace(',', '.');
  }
  
  return parseFloat(clean);
}

/**
 * Generates quick trading advice with AI Fallback.
 */
function generateSniperAI(alerts) {
  const API_KEY = "TUO_API_KEY_QUI"; // REPLACE THIS WITH YOUR ACTUAL API KEY or use PropertiesService
  const MODELS = ["gemini-2.0-flash", "gemini-flash-latest", "gemini-2.0-flash-lite"];

  const prompt = `
    ROLE: You are an Expert Day Trader.
    EVENTS: The following assets are moving fast today:
    ${JSON.stringify(alerts)}
    
    TASK:
    Provide immediate, short advice (In Italian).
    If dropping: "Buy the dip or catching a falling knife?"
    If pumping: "Take profit or let it ride?"
    Max 2 sentences.
  `;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  for (let i = 0; i < MODELS.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[i]}:generateContent?key=${API_KEY}`;
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
      }
    } catch(e) { console.warn(`Sniper AI Model ${MODELS[i]} failed.`); }
  }
  return "Market is volatile. Trade with caution.";
}