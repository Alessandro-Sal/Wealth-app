/**
 * App_SniperAlert.js
 * Scans LIVE portfolio for rapid price movements (> 4% or < -4%).
 * - Stocks/ETFs: Uses the "pct" (Percentage Change) column from the sheet.
 * - Crypto: Calculates volatility based on the difference vs. the LAST SCRIPT RUN.
 * Uses ROBUST AI for market commentary.
 * * SCHEDULE CONSTRAINT: Runs ONLY on Weekdays (Mon-Fri) between 16:00 and 22:30.
 */

function runMarketSniper() {
  const now = new Date();
  const day = now.getDay();   // 0 = Sunday, 6 = Saturday
  const hour = now.getHours();
  const minute = now.getMinutes();

  // --- 1. TIME & DAY CONSTRAINT CHECK ---
  
  // A. Day Check: Run ONLY on Weekdays (Monday to Friday).
  // If it's a Weekend (0 or 6), STOP.
  if (day === 0 || day === 6) {
    console.log("Sniper Skipped: Weekend (Sat-Sun).");
    return;
  }

  // B. Time Check: Run ONLY between 16:00 and 22:30.
  // Before 16:00? STOP.
  if (hour < 16) {
    console.log("Sniper Skipped: Too early (Before 16:00).");
    return;
  }
  
  // After 22:30? STOP.
  if (hour > 22 || (hour === 22 && minute >= 30)) {
    console.log("Sniper Skipped: Too late (After 22:30).");
    return;
  }

  console.log(`Sniper Active: Weekday ${now.toLocaleTimeString()}`);

  // --- 2. Get Live Portfolio Data ---
  let portfolio;
  try {
     portfolio = getLivePortfolio(); 
  } catch(e) {
     console.error("Sniper Error: Could not fetch portfolio.");
     return;
  }

  // Combine Stocks, ETFs, and Crypto into one array
  const allAssets = [...portfolio.stocks, ...portfolio.etfs, ...portfolio.crypto];
  
  // Retrieve stored prices from the last run
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastPrices = scriptProperties.getProperties();
  
  let alerts = [];
  let updates = {}; // Object to store current prices for the next run

  // --- 3. Scan for Volatility ---
  allAssets.forEach(asset => {
    let changeVal = 0.0;
    let changeStr = "0%";
    
    // Check if asset has a valid ticker/name
    if (!asset.t) return;

    // --- CASE A: STOCKS & ETFS (Use 'pct' - Percentage Change) ---
    // We check if 'pct' exists and is not a placeholder.
    if (asset.pct && asset.pct !== "" && asset.pct !== "0%" && asset.sector !== "Crypto") {
       changeStr = asset.pct;
       // Clean string: "-6,00%" -> -6.00 (replace comma with dot for JS math)
       changeVal = parseFloat(String(changeStr).replace('%', '').replace(',', '.'));
    } 
    // --- CASE B: ASSETS WITHOUT % (e.g., Crypto or missing data) ---
    // We calculate volatility manually by comparing Current Price vs. Last Run Price
    else {
       // Helper to clean price string
       const cleanPrice = parsePrice(asset.price || asset.val); // Fallback to val if price missing
       
       if (cleanPrice > 0) {
           // Create a unique cache key (e.g., LAST_PRICE_BTC)
           const cacheKey = "LAST_PRICE_" + asset.t.replace(/\s/g, ''); 
           
           // Get previous price (if any)
           const lastPriceStr = lastPrices[cacheKey];
           const lastPrice = lastPriceStr ? parseFloat(lastPriceStr) : 0;

           // Only calculate if we have a valid previous price to compare against
           if (lastPrice > 0) {
             changeVal = ((cleanPrice - lastPrice) / lastPrice) * 100;
             changeStr = changeVal.toFixed(2).replace('.', ',') + "% (vs Last Run)";
           }

           // Store the current price to be used as "lastPrice" in the next run
           updates[cacheKey] = String(cleanPrice);
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

  // Save the updated Crypto prices to script properties
  if (Object.keys(updates).length > 0) {
    scriptProperties.setProperties(updates);
  }

  // --- 4. Send Alert if Triggered ---
  if (alerts.length > 0) {
    const aiComment = generateSniperAI(alerts);
    
    MailApp.sendEmail({
      to: "alessandro.saladino01@gmail.com",
      subject: `🔥 MARKET MOVER: ${alerts.length} Asset in Motion`,
      htmlBody: `
      <div style="font-family: Arial, sans-serif; padding: 20px; border: 2px solid #e74c3c; border-radius: 8px; max-width: 600px;">
        <h2 style="color: #c0392b; margin-top: 0;">🎯 Sniper Alert</h2>
        <p>Volatile movements detected during weekday session:</p>
        <ul style="font-size: 16px; background-color: #fff0f0; padding: 15px; border-radius: 5px;">
          ${alerts.map(a => `<li style="margin-bottom: 8px; list-style-type: none;">${a}</li>`).join('')}
        </ul>
        <hr style="border: 0; border-top: 1px solid #ccc; margin: 20px 0;">
        <p><strong>🧠 AI Analyst Opinion:</strong></p>
        <p style="font-style: italic; color: #555;">${aiComment}</p>
        <br>
        <div style="font-size: 11px; color: #999; text-align: center;">
          *Crypto change is calculated based on movement since the last script run.<br>
          Generated automatically by Wealth App Sniper.
        </div>
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
  if (clean.includes(',') && clean.includes('.')) {
      // Assume Italian format (1.000,00) -> remove dots, replace comma with dot
      clean = clean.replace(/\./g, '').replace(',', '.');
  } else if (clean.includes(',')) {
      // If only comma exists (e.g. 12,50), replace with dot
      clean = clean.replace(',', '.');
  }
  
  return parseFloat(clean) || 0;
}

/**
 * Generates quick trading advice with AI Fallback and REAL-TIME WEB SEARCH.
 */
function generateSniperAI(alerts) {
  const API_KEY = GEMINI_API_KEY; 
  
  if (!API_KEY) {
      console.warn("Sniper AI: GEMINI_API_KEY global variable not found.");
      return "AI Commentary Unavailable (Missing Key).";
  }

  // Updated Fallback based on available API models, prioritizing Lite/Flash versions
  // to avoid 429 Rate Limit errors while keeping web search capabilities.
  const MODELS = [
    "gemini-2.5-flash-lite", 
    "gemini-2.0-flash-lite", 
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-3-flash-preview"
  ];

  const prompt = `
    ROLE: You are an Expert Crypto & Swing Trader.
    CONTEXT: Weekday market session. The following assets have triggered high-volatility alerts:
    ${JSON.stringify(alerts)}
    
    TASK:
    1. Search the web for the latest breaking news TODAY regarding these specific assets.
    2. Provide a short explanation (In Italian) of WHY they are moving.
    3. You MUST include the name of the news source or the title of the article you found.
    4. Conclude with a quick actionable advice. 
    Keep it punchy. Max 4 sentences total. language: Italian. Use emojis if you want, but keep it professional.
  `;

  const payload = { 
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ google_search: {} }] 
  };
  
  const options = { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify(payload), 
    muteHttpExceptions: true 
  };

  for (let i = 0; i < MODELS.length; i++) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[i]}:generateContent?key=${API_KEY}`;
    const res = UrlFetchApp.fetch(url, options);
    const responseCode = res.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(res.getContentText());
      
      // Debugging: Verify if Grounding (Web Search) was actually used
      const metadata = data.candidates[0].groundingMetadata;
      if (metadata && metadata.webSearchQueries) {
          console.log(`Sniper AI performed web searches: ${metadata.webSearchQueries.join(", ")}`);
      } else {
          console.log("Sniper AI Warning: Responded without triggering Google Search.");
      }

      let text = data.candidates[0].content.parts[0].text;
      return text.replace(/\*/g, ''); 
    } else {
      console.warn(`Model ${MODELS[i]} failed. Code: ${responseCode}`);
      
      // If we hit a Rate Limit (429), pause for 10 seconds before trying the next model
      if (responseCode === 429) {
        console.log(`Rate limit hit on ${MODELS[i]}. Waiting 10 seconds before next attempt...`);
        Utilities.sleep(10000); 
      }
    }
  }
  
  return "Market is volatile today. AI quota exceeded, trade with caution.";
}