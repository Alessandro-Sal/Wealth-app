/**
 * App_SniperAlert.js
 * Scans LIVE portfolio for rapid price movements (> 4% or < -4%).
 * - Stocks/ETFs: Uses the "pct" (Percentage Change) column. Evaluates per-asset market hours.
 * - Crypto: Calculates rolling 24h volatility using script cache. Runs 24/7.
 * Uses ROBUST AI for market commentary via Universal Router (OpenRouter -> Gemini).
 */

function runMarketSniper() {
  const now = new Date();
  const nowMs = now.getTime();
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  console.log(`Sniper Scan Started: ${now.toLocaleTimeString()}`);

  // --- 1. Get Live Portfolio Data ---
  let portfolio;
  try {
     portfolio = getLivePortfolio(); 
  } catch(e) {
     console.error("Sniper Error: Could not fetch portfolio.");
     return;
  }

  // Tag crypto assets explicitly so they are always recognized, 
  // regardless of missing sheet columns.
  const taggedCrypto = portfolio.crypto.map(c => {
      c._isCryptoAsset = true;
      return c;
  });

  // Combine Stocks, ETFs, and explicitly tagged Crypto into one array
  const allAssets = [...portfolio.stocks, ...portfolio.etfs, ...taggedCrypto];
  
  // Retrieve stored prices from the last runs
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastPrices = scriptProperties.getProperties();
  
  let alerts = [];
  let updates = {}; // Object to store current histories for the next run
  
  let traditionalMarketsScanned = 0;
  let cryptoScanned = 0;

  // --- 2. Scan for Volatility ---
  allAssets.forEach(asset => {
    let changeVal = 0.0;
    let changeStr = "0%";
    
    // Check if asset has a valid ticker/name
    if (!asset.t) return;

    // Check if asset is crypto by sector, exchange, or our explicit tag
    const isCrypto = (asset.sector === "Crypto" || String(asset.ex).toUpperCase() === 'CRYPTO' || asset._isCryptoAsset);

    // --- MARKET HOURS CHECK (For Stocks & ETFs only) ---
    if (!isCrypto) {
        // Use the server-side logic to check if the specific exchange is open
        const marketStatus = getMarketStateServer(asset.ex);
        
        // If the market is completely closed (not PRE_MARKET, OPEN, or CLOSING_SOON), skip this asset.
        if (marketStatus === 'CLOSED') {
            return; 
        }
        traditionalMarketsScanned++;
    } else {
        cryptoScanned++;
    }

    // --- CASE A: STOCKS & ETFS (Use 'pct' - Percentage Change) ---
    // We check if 'pct' exists and is not a placeholder.
    if (!isCrypto && asset.pct && asset.pct !== "" && asset.pct !== "0%") {
       changeStr = asset.pct;
       // Clean string: "-6,00%" -> -6.00 (replace comma with dot for JS math)
       changeVal = parseFloat(String(changeStr).replace('%', '').replace(',', '.'));
    } 
    // --- CASE B: CRYPTO or ASSETS WITHOUT % ---
    // We calculate 24h volatility manually by storing a history queue in properties
    else {
       // Helper to clean price string
       const cleanPrice = parsePrice(asset.price || asset.val); // Fallback to val if price missing
       
       if (cleanPrice > 0) {
           // Create a unique cache key for the 24h history
           const cacheKey = "HISTORY_24H_" + asset.t.replace(/\s/g, ''); 
           
           let history = [];
           const historyStr = lastPrices[cacheKey];
           if (historyStr) {
               try {
                   history = JSON.parse(historyStr);
               } catch(e) {
                   history = [];
               }
           }

           // Add current price and timestamp to history
           history.push({ ts: nowMs, p: cleanPrice });

           // Filter out entries older than 24 hours
           history = history.filter(entry => entry.ts >= (nowMs - ONE_DAY_MS));

           // Calculate variation compared to the oldest price in our 24h window
           if (history.length > 1) {
             const oldestPrice = history[0].p;
             changeVal = ((cleanPrice - oldestPrice) / oldestPrice) * 100;
             changeStr = changeVal.toFixed(2).replace('.', ',') + "% (24h)";
           }

           // Store the updated history array back as a JSON string
           updates[cacheKey] = JSON.stringify(history);
       }
    }

    // THRESHOLDS: Drop < -4% or Pump > +5%
    if (!isNaN(changeVal)) {
        if (changeVal <= -4.0) {
          alerts.push(`🔻 <b>${asset.t}</b> is crashing: <span style="color:red">${changeStr}</span>`);
        } else if (changeVal >= 5.0) {
          alerts.push(`🚀 <b>${asset.t}</b> is pumping: <span style="color:green">+${changeStr}</span>`);
        }
    }
  });

  // Save the updated Crypto/Manual price histories to script properties
  if (Object.keys(updates).length > 0) {
    scriptProperties.setProperties(updates);
  }

  // --- 3. Send Alert if Triggered ---
  if (alerts.length > 0) {
    const aiComment = generateSniperAI(alerts);
    
    // Determine session type for the email subject/header
    let sessionType = "Crypto-Only Session";
    if (traditionalMarketsScanned > 0) {
        sessionType = "Active Market Session";
    }
    
    MailApp.sendEmail({
      to: "alessandro.saladino01@gmail.com",
      subject: `🔥 MARKET MOVER: ${alerts.length} Asset in Motion`,
      htmlBody: `
      <div style="font-family: Arial, sans-serif; padding: 20px; border: 2px solid #e74c3c; border-radius: 8px; max-width: 600px;">
        <h2 style="color: #c0392b; margin-top: 0;">🎯 Sniper Alert</h2>
        <p>Volatile movements detected during <b>${sessionType}</b>:</p>
        <ul style="font-size: 16px; background-color: #fff0f0; padding: 15px; border-radius: 5px;">
          ${alerts.map(a => `<li style="margin-bottom: 8px; list-style-type: none;">${a}</li>`).join('')}
        </ul>
        <hr style="border: 0; border-top: 1px solid #ccc; margin: 20px 0;">
        <p><strong>🧠 AI Analyst Opinion:</strong></p>
        <p style="font-style: italic; color: #555;">${aiComment}</p>
        <br>
        <div style="font-size: 11px; color: #999; text-align: center;">
          *Crypto change is calculated based on movement over a rolling 24-hour period.<br>
          Generated automatically by Wealth App Sniper.
        </div>
      </div>`
    });
    console.log(`Sniper Alert Sent: ${alerts.length} assets. (Scanned: ${traditionalMarketsScanned} Trad, ${cryptoScanned} Crypto)`);
  } else {
    console.log(`Sniper Scan: Markets calm. (Scanned: ${traditionalMarketsScanned} Trad, ${cryptoScanned} Crypto)`);
  }
}

/**
 * Server-side adaptation of getMarketState.
 * Evaluates if a specific exchange is open right now based on Time Zones.
 */
function getMarketStateServer(exchangeCode) {
  if (!exchangeCode || String(exchangeCode).toUpperCase() === 'CRYPTO') return 'OPEN';
  
  const now = new Date(); 
  let timeZone = "UTC"; 
  let openH = 9, openM = 0, closeH = 17, closeM = 0;
  
  switch (String(exchangeCode).toUpperCase()) {
      case 'US': case 'USA': case 'NASDAQ': case 'NYSE': case 'AMEX': 
          timeZone = "America/New_York"; openH = 9; openM = 30; closeH = 16; closeM = 0; break;
      case 'IT': case 'ITA': case 'MIL': case 'MTA': case 'DE': case 'GER': case 'XETRA': case 'EU': case 'EAM': case 'TDG': case 'CPH': 
          timeZone = "Europe/Berlin"; openH = 9; openM = 0; closeH = 17; closeM = 30; break;
      case 'UK': case 'LSE': 
          timeZone = "Europe/London"; openH = 8; openM = 0; closeH = 16; closeM = 30; break;
      default: 
          return 'CLOSED';
  }
  
  try {
      // Create a date string localized to the specific exchange's timezone
      const localTimeStr = now.toLocaleString("en-US", {timeZone: timeZone}); 
      const localDate = new Date(localTimeStr); 
      
      const day = localDate.getDay(); 
      const h = localDate.getHours(); 
      const m = localDate.getMinutes();
      
      const curMins = h * 60 + m; 
      const openMins = openH * 60 + openM; 
      const closeMins = closeH * 60 + closeM;
      
      // Weekend check (local to the exchange)
      if (day === 0 || day === 6) return 'CLOSED';
      
      // Inside market hours
      if (curMins >= openMins && curMins < closeMins) { 
          if (closeMins - curMins <= 60) return 'CLOSING_SOON'; 
          return 'OPEN'; 
      }
      
      // Pre-market (1 hour before)
      if (curMins >= (openMins - 60) && curMins < openMins) return 'PRE_MARKET';
      
      return 'CLOSED';
  } catch (e) { 
      console.warn(`Timezone calculation error for ${exchangeCode}: ${e}`);
      return 'CLOSED'; 
  }
}

/**
 * Helper function to parse currency strings.
 */
function parsePrice(priceStr) {
  if (!priceStr) return 0;
  let clean = String(priceStr).replace(/[^\d,\.-]/g, ''); 
  if (clean.includes(',') && clean.includes('.')) {
      clean = clean.replace(/\./g, '').replace(',', '.');
  } else if (clean.includes(',')) {
      clean = clean.replace(',', '.');
  }
  return parseFloat(clean) || 0;
}

/**
 * Generates quick trading advice using a cascading AI fallback system.
 * Tries OpenRouter first (for Claude/GPT), falls back to Gemini if it fails.
 */
function generateSniperAI(alerts) {
  const prompt = `
    ROLE: You are an Expert Crypto & Swing Trader.
    CONTEXT: Live market session. The following assets have triggered high-volatility alerts:
    ${JSON.stringify(alerts)}
    
    TASK:
    1. Provide a short explanation (In Italian) of WHY they might be moving.
    2. Conclude with a quick actionable advice. 
    Keep it punchy. Max 4 sentences total. language: Italian. Use emojis if you want, but keep it professional.
  `;

  let response = null;

  // PLAN A: Try OpenRouter (Claude 3.5 or GPT-4o)
  console.log("Sniper AI: Attempting OPENROUTER...");
  response = fetchUniversalAI(prompt, 'OPENROUTER');

  // PLAN B: Fallback to Gemini (Native cascade)
  if (!response) {
    console.log("Sniper AI: OPENROUTER failed. Falling back to GEMINI...");
    response = fetchUniversalAI(prompt, 'GEMINI', true);
  }

  // PLAN C: Ultimate Fallback
  if (!response) {
    return "Market is volatile today. AI services are currently overloaded, trade with caution.";
  }

  return response.replace(/\*/g, ''); // Clean markdown for email
}