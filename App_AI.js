/**
 * Main chat interface for the AI Assistant.
 * Handles user authentication via PIN to inject financial context (Net Worth, Savings).
 * Implements a robust fallback strategy cycling through multiple Gemini models (Flash 2.0, Latest, Lite) to ensure high availability.
 * * @param {string} userQuestion - The user's query.
 * @param {string} sessionPin - The session PIN to validate access to sensitive data.
 * @param {string} historyJson - Previous chat history context.
 * @return {string} The AI response.
 */
function askGemini(userQuestion, sessionPin, historyJson) {
  const SECRET_PIN = PropertiesService.getScriptProperties().getProperty('APP_PIN');
  
  // Reference to the global key (ensure GEMINI_API_KEY is defined in Secrets.gs or Global)
  const API_KEY = GEMINI_API_KEY; 

  const MODELS = [
    "gemini-2.0-flash",      // Primary: Fastest and most modern
    "gemini-flash-latest",   // Fallback 1: Latest stable version
    "gemini-2.0-flash-lite"  // Fallback 2: Lightweight version
  ];

  const isAuthorized = (String(sessionPin).trim() === SECRET_PIN);
  
  let messages = [];
  let systemPrompt = "";

  if (isAuthorized) {
    const dash = getDashboardData();
    const savings = getMonthlySavings();
    
    // Prompt translated to English, but keeps the context clear
    systemPrompt = `You are Wealth AI. User VERIFIED.
    Data: Net Worth ${dash.totalNetWorth}, Liquid ${dash.liquidNetWorth}.
    Current Month: In ${savings.income}, Out ${savings.expenses}.
    Keep responses concise.`;
  } else {
    systemPrompt = "You are Wealth AI. User GUEST. Do not reveal any financial data.";
  }

  messages.push({ role: "user", parts: [{ text: systemPrompt }] });
  
  // Load chat history if available
  if (historyJson) {
    try {
      const prevChat = JSON.parse(historyJson);
      if(Array.isArray(prevChat)) prevChat.forEach(msg => messages.push(msg));
    } catch(e) {}
  }
  messages.push({ role: "user", parts: [{ text: userQuestion }] });

  const payload = { contents: messages };
  const options = {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  // --- MODEL RETRY LOOP ---
  for (let m = 0; m < MODELS.length; m++) {
    const modelName = MODELS[m];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${API_KEY}`;
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const json = JSON.parse(response.getContentText());

      if (code === 200 && !json.error) {
        return json.candidates[0].content.parts[0].text;
      }
      
      // If error is not 429 (Rate Limit), try the next model
      if (code !== 429) {
        console.warn(`Model ${modelName} failed (${code}). Trying next.`);
        continue;
      }
      
    } catch (e) {
      console.error(e);
    }
  }

  return "All AI models are currently busy. Please try again shortly.";
}


/**
 * Performs a deep-dive risk analysis of the portfolio.
 * Aggregates liquid assets, stocks, and crypto to generate a "Stress Test" report.
 * Uses aggressive caching (12 hours) to minimize token usage for this heavy operation.
 * * @return {Object} JSON object containing risk score, concentration analysis, and market crash simulations.
 */
function getPortfolioRiskAnalysis() {
  const CACHE_KEY = "GEMINI_RISK_DEEP_V1"; 
  const cache = CacheService.getScriptCache();
  
  // Check cache first (Analysis is heavy)
  const cachedResult = cache.get(CACHE_KEY);
  if (cachedResult) return JSON.parse(cachedResult);

  // 1. Retrieve Macro Data
  const dash = getDashboardData();
  
  // 2. Retrieve Stock & Crypto details (Top 15 positions to save context window)
  const portfolio = getLivePortfolio();
  
  // Helper to clean data for AI
  const cleanList = (list) => list.slice(0, 15).map(i => `${i.t} (${i.pct})`).join(", ");
  const topStocks = cleanList(portfolio.stocks);
  const topCrypto = cleanList(portfolio.crypto);

  // 3. Construct Advanced Prompt (Instructions in English, Output language requested: Italian)
  const prompt = `
    You are a Senior Financial Risk Manager. Analyze this personal portfolio:
    
    ASSET ALLOCATION:
    - Cash: ${dash.summary.cash.percent} (Total Liquid: ${dash.liquidNetWorth})
    - Stocks: ${dash.summary.stocks.percent}
    - Crypto: ${dash.summary.crypto.percent}
    - ETFs: ${dash.summary.etfs.percent}
    
    TOP HOLDINGS (Ticker and Performance):
    - Stocks: ${topStocks}
    - Crypto: ${topCrypto}

    Provide a STRICT JSON output with this structure:
    {
      "riskScore": "1-100 (1=Safe, 100=Gambling)",
      "riskLevel": "Low/Medium/High/Extreme",
      "summary": "Synthetic analysis in 1 sentence",
      "concentration": "Concentration analysis (e.g., Too much Tech, Too much USA)",
      "stressTest": {
        "marketCrash": "Estimated loss % if S&P500 drops 20%",
        "cryptoWinter": "Estimated loss % if Bitcoin drops 50%"
      },
      "suggestions": ["Practical advice 1", "Practical advice 2", "Practical advice 3"]
    }
    
    Be brutal and honest.
  `;

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = [
    "gemini-2.0-flash", 
    "gemini-flash-latest",
    "gemini-2.0-flash-lite"
  ];
  
  let finalResult = { riskLevel: "N/A", summary: "AI unavailable" };

  for (let m = 0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const res = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        // Clean Markdown JSON
        text = text.replace(/```json/g, "").replace(/```/g, "").trim();
        finalResult = JSON.parse(text);
        
        // Save to cache
        cache.put(CACHE_KEY, JSON.stringify(finalResult), 86400); // 12 hours
        break; 
      }
    } catch (e) {
      console.error("Risk AI Error: " + e.toString());
    }
  }

  return finalResult;
}

/**
 * Utility function to clear specific server-side cache keys.
 * Used to force a refresh of the Risk Analysis or Crypto Sentiment data.
 */
function clearServerCache() {
  const cache = CacheService.getScriptCache();
  // Added new key "GEMINI_RISK_DEEP_V1" to the clearing list
  cache.removeAll(['GEMINI_RISK_ANALYSIS_PCT_V2', 'GEMINI_RISK_DEEP_V1', 'CRYPTO_FNG']); 
  return "Server Cache Cleared";
}


/**
 * AI-powered parser for expenses (Smart Input).
 * Converts unstructured voice text or receipt images into structured JSON data.
 * Recognizes categories, amounts, and dates automatically.
 * * @param {string} inputData - The text prompt or base64 image string.
 * @param {string} mode - The input mode: 'voice' or 'image'.
 * @return {Object} Structured expense object {type, category, amount, desc, date}.
 */
function parseExpenseAI(inputData, mode) {
  const API_KEY = GEMINI_API_KEY; 
  
  // Model Sequence (Try in order)
  const MODELS = [
    "gemini-2.0-flash", 
    "gemini-flash-latest",
    "gemini-2.0-flash-lite"
  ];

  // ITALIAN CATEGORIES PRESERVED AS REQUESTED
  const CATS = "Alimentazione, Alloggio, Trasporti, Free-Time, Necessità, Regali, Uscite, Viaggi, Altro, Stipendio";
  
  let userContent = [];
  
  // Prompt Construction (Instructions in English, but forcing Italian Categories)
  let systemText = `You are an accounting assistant. Analyze the input and extract transaction data.
  Allowed Categories (Must use one of these exact strings): [${CATS}].
  Today is: ${new Date().toLocaleDateString()}.
  
  Expected Output strictly valid JSON (no markdown, no backticks):
  {
    "type": "Expense" or "Income",
    "category": "One of the allowed categories above or 'Altro'",
    "amount": 0.00 (use dot for decimals, pure number),
    "desc": "Short description (e.g., 'Pizza at Michele')",
    "date": "YYYY-MM-DD" (if not specified, use today)
  }`;

  if (mode === 'voice') {
    userContent.push({ text: `Analyze this voice text: "${inputData}"` });
  } else if (mode === 'image') {
    userContent.push({ text: "Analyze this receipt/invoice image." });
    userContent.push({ inline_data: { mime_type: "image/jpeg", data: inputData } });
  }

  const payload = {
    contents: [{ role: "user", parts: [{ text: systemText }, ...userContent] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // --- MODEL LOOP ---
  let lastError = "";

  for (let i = 0; i < MODELS.length; i++) {
    const modelName = MODELS[i];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${API_KEY}`;
    
    try {
      console.log(`Attempting model: ${modelName}`); 
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const textRaw = response.getContentText();

      if (code !== 200) {
        console.warn(`Model ${modelName} failed with code ${code}: ${textRaw}`);
        lastError = `API Error (${code})`;
        continue; 
      }

      const json = JSON.parse(textRaw);

      if (!json.candidates || json.candidates.length === 0) {
        console.warn(`Model ${modelName} returned no candidates.`);
        lastError = "No AI result generated.";
        continue;
      }

      // --- SUCCESS ---
      let cleanText = json.candidates[0].content.parts[0].text;
      cleanText = cleanText.replace(/```json/g, "").replace(/```/g, "").trim();
      
      return JSON.parse(cleanText);

    } catch (e) {
      console.error(`Exception on ${modelName}: ${e.toString()}`);
      lastError = e.toString();
    }
  }

  return { error: "All AI models are busy or failed. Last error: " + lastError };
}

/**
 * Generates a comprehensive "Hedge Fund" style market report.
 * 1. Scrapes Macro data (S&P500, VIX, US10Y).
 * 2. Calculates weighted portfolio metrics (Beta, P/E) excluding Crypto.
 * 3. Asks AI to synthesize sentiment, technical analysis, and upcoming market events.
 * * @return {Object} Integrated object with macro data, portfolio metrics, and AI commentary.
 */
function getMarketInsightsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 1. MACRO DATA ---
  let macro = { 
    spx: 0, dow: 0, nasdaq: 0, russell: 0, vix: 0, us10y: 0, 
    cryptoTrend: "N/A" 
  };
  
  const cleanPct = (valStr) => {
    if (!valStr) return 0;
    let s = String(valStr).replace(/[€$£%\s]/g, '').trim().replace(',', '.');
    return parseFloat(s) || 0;
  };

  try {
    const sheet = ss.getSheetByName("Stock Market Dashboard");
    if (sheet) {
      macro.me = cleanPct(sheet.getRange("B7").getDisplayValue());
      const indices = sheet.getRange("C5:G5").getDisplayValues()[0];
      macro.spx = cleanPct(indices[0]);
      macro.dow = cleanPct(indices[1]); 
      macro.nasdaq = cleanPct(indices[2]);
      macro.russell = cleanPct(indices[3]);
      macro.vix = cleanPct(indices[4]);
      macro.us10y = cleanPct(sheet.getRange("E12").getDisplayValue());
    }
  } catch(e) { console.error("Error Dashboard Data: " + e); }

  // --- 2. EXTENDED DATA RETRIEVAL ---
  const port = getLivePortfolio(); 
  
  let cryptoTotalChange = 0;
  let cryptoCount = 0;
  if (port.crypto && port.crypto.length > 0) {
    port.crypto.forEach(c => {
      if (c.pct) { 
        cryptoTotalChange += cleanPct(c.pct);
        cryptoCount++;
      }
    });
    const avgCrypto = cryptoCount > 0 ? (cryptoTotalChange / cryptoCount).toFixed(2) : "0";
    macro.cryptoTrend = `${avgCrypto}% (Basket of ${port.crypto.length} coins)`;
  }

  const equityAssets = [...port.stocks, ...port.etfs];

  let totalEquityVal = 0;
  let weightedBeta = 0;
  let weightedPE = 0;
  let peEligibleVal = 0;
  let sectorMap = {};
  let countryMap = {};

  const parseVal = (v) => parseFloat(String(v).replace(/[^\d.-]/g, '').replace(',', '.')) || 0;

  const enrichedEquity = equityAssets.map(a => {
    const val = parseVal(a.val);
    const beta = parseFloat(a.beta) || 1; 
    const pe = parseFloat(a.pe) || 0;
    const price = parseVal(a.price);
    const yearH = parseVal(a.yearH);
    const yearL = parseVal(a.yearL);
    
    let rangePos = "N/A";
    if (yearH > yearL && price > 0) {
       rangePos = ((price - yearL) / (yearH - yearL) * 100).toFixed(0) + "%";
    }

    totalEquityVal += val;
    weightedBeta += (beta * val);
    
    if (pe > 0 && a.sector !== 'ETFs') {
      weightedPE += (pe * val);
      peEligibleVal += val;
    }

    const s = a.sector || "Other";
    sectorMap[s] = (sectorMap[s] || 0) + val;
    const c = a.ctry || "Global";
    countryMap[c] = (countryMap[c] || 0) + val;

    return {
      t: a.t,
      val: val,
      dCh: a.dCh,
      details: {
        pe: pe,
        beta: beta,
        mkt: a.mkt,
        eps: a.eps,
        vol: a.vol,
        avgVol: a.avgVol,
        rangePos: rangePos,
        ind: a.ind,
        ctry: a.ctry
      }
    };
  }).sort((a, b) => b.val - a.val); 

  // --- 3. FINAL METRICS ---
  const portBeta = totalEquityVal > 0 ? (weightedBeta / totalEquityVal).toFixed(2) : 1;
  const portPE = peEligibleVal > 0 ? (weightedPE / peEligibleVal).toFixed(1) : "N/A";

  const topSectors = Object.entries(sectorMap)
    .sort((a,b) => b[1]-a[1]).slice(0,5)
    .map(([k,v]) => `${k}:${((v/totalEquityVal)*100).toFixed(0)}%`).join(", ");
    
  const topCountries = Object.entries(countryMap)
    .sort((a,b) => b[1]-a[1]).slice(0,3)
    .map(([k,v]) => k).join(", ");

  const assetsStr = enrichedEquity
    .slice(0, 30)
    .map(a => {
      const d = a.details;
      let volStatus = "";
      const vNow = parseFloat(d.vol);
      const vAvg = parseFloat(d.avgVol);
      if(vNow && vAvg) {
        if (vNow > vAvg * 1.5) volStatus = "| HighVol";
        else if (vNow < vAvg * 0.5) volStatus = "| LowVol";
      }
      return `${a.t} (${a.dCh} | β:${d.beta} | PE:${d.pe>0?d.pe:'-'} | 52wPos:${d.rangePos} ${volStatus})`;
    })
    .join("\n    ");

  const myTickers = enrichedEquity.map(a => a.t).join(", ");

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  const today = new Date().toISOString().split('T')[0];

  // Prompt completely in English
  const prompt = `
    Role: Algorithmic Hedge Fund Manager. Date: ${today}.
    
    === 1. GLOBAL MACRO PICTURE ===
    - USA Indices: S&P500 ${macro.spx}% | Dow Jones ${macro.dow}% | Nasdaq ${macro.nasdaq}% | Russell ${macro.russell}%
    - Risk Meters: VIX ${macro.vix} | US 10Y Yield ${macro.us10y}%
    - Crypto Sentiment (Risk-On Indicator): ${macro.cryptoTrend}

    === 2. EQUITY PORTFOLIO ANALYSIS (Stocks & ETFs - No Crypto) ===
    - Daily Performance: ${macro.me}%
    - FUNDAMENTAL METRICS (Weighted):
      * Beta: ${portBeta} ( >1 Aggressive, <1 Defensive)
      * P/E Ratio: ${portPE} (Value vs Growth. Ref SP500 ~21)
      * Geo Exposure: ${topCountries}
      * Sector Exposure: ${topSectors}
    
    - DETAILED TOP HOLDINGS (Ticker | Day% | Beta | P/E | Range Position 0-100% | Volume):
    [
    ${assetsStr}
    ]

    === 3. TASK (Deep Reasoning) ===
    A. **Macro & Crypto Analysis:** How do Yields (US10Y), VIX, and Crypto sentiment interact today? Risk-On or Flight to Safety?
    B. **Portfolio Diagnostics:**
       - Style: Do Beta (${portBeta}) and P/E (${portPE}) explain today's performance?
       - Technicals: Note if top holdings are near highs (52wPos > 90%) or oversold. Mention volume anomalies.
    C. **Sentiment Score:** 0 (Panic) to 10 (Euphoria).
    D. **Events:**
       - MARKET: 5 relevant upcoming macro events.
       - PORTFOLIO: Earnings or Dividends expected ONLY for: [${myTickers}].

    JSON STRICT OUTPUT:
    {
      "sentiment_score": 5.5,
      "sentiment_label": "Fear/Neutral/Greed",
      "macro_analysis": "Integrated analysis...",
      "portfolio_analysis": "Technical and fundamental analysis...",
      "market_events": [{"ticker": "MACRO", "event": "Event", "date": "YYYY-MM-DD"}],
      "portfolio_events": [{"ticker": "TICKER", "event": "Earnings/Div", "date": "YYYY-MM-DD"}]
    }
  `;

  let aiRes = { 
    sentiment_score: 5, sentiment_label: "Loading", 
    macro_analysis: "...", portfolio_analysis: "...", 
    market_events: [], portfolio_events: [] 
  };

  for (let m=0; m<MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
      
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        let txt = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        txt = txt.replace(/```json/g, "").replace(/```/g, "").trim();
        const s = txt.indexOf('{'); const e = txt.lastIndexOf('}');
        if (s !== -1 && e !== -1) aiRes = JSON.parse(txt.substring(s, e + 1));
        break;
      }
    } catch(err) { console.warn("AI Err: " + err); }
  }

  return {
    ...macro,
    metrics: { beta: portBeta, pe: portPE },
    sentiment: { score: aiRes.sentiment_score, label: aiRes.sentiment_label },
    analysis: { macro: aiRes.macro_analysis, portfolio: aiRes.portfolio_analysis },
    market_events: aiRes.market_events,
    portfolio_events: aiRes.portfolio_events
  };
}

/**
 * estimates annual dividend income using AI.
 * Cleans European number formats, fetches current dividend yields/payment months via Gemini,
 * and calculates the total projected yearly return for the top 15 assets.
 * * @return {Object} Total yearly estimate and a detailed list of paying assets.
 */
function fetchDividendData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const port = getLivePortfolio();
  const assets = [...port.stocks, ...port.etfs];
  
  const parseValue = (val) => {
    if (!val) return 0;
    let s = String(val).replace(/[€$£\s%]/g, ''); 
    if (s.includes('.') && s.includes(',')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else if (s.includes(',')) {
      s = s.replace(',', '.');
    }
    return parseFloat(s) || 0;
  };

  const assetList = assets
    .map(a => ({
       ticker: a.t.toUpperCase(), 
       value: parseValue(a.val)   
    }))
    .filter(a => a.value > 0) 
    .sort((a, b) => b.value - a.value)
    .slice(0, 15);

  if (assetList.length === 0) return { totalYearly: 0, items: [] };

  const assetString = assetList.map(a => a.ticker).join(", ");

  const API_KEY = GEMINI_API_KEY; 
  
  const MODELS = [
    "gemini-2.0-flash",    
    "gemini-1.5-flash",    
    "gemini-flash-latest"
  ];
  
  // English Prompt
  const prompt = `
    I have these assets: [${assetString}].
    
    For each one, estimate for TODAY:
    1. The current annual "Dividend Yield" in % (e.g., 0.03 for 3%). If no dividend, 0.
    2. The expected month for the NEXT payment (e.g., "Mar", "Apr").
    
    Reply ONLY with a JSON Array:
    [{"t": "TICKER", "y": 0.034, "m": "Mar"}]
  `;

  let aiData = [];

  for (let m=0; m<MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const options = { 
        method: "post", 
        contentType: "application/json", 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      };
      
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        let txt = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        txt = txt.replace(/```json/g, "").replace(/```/g, "").trim();
        
        const firstBracket = txt.indexOf('[');
        const lastBracket = txt.lastIndexOf(']');
        if (firstBracket !== -1 && lastBracket !== -1) {
            txt = txt.substring(firstBracket, lastBracket + 1);
            aiData = JSON.parse(txt);
            break; 
        }
      }
    } catch(e) { console.warn("AI Model Error:", e); }
  }

  let totalYearly = 0;
  let finalItems = [];

  assetList.forEach(myAsset => {
      const info = aiData.find(d => d.t.toUpperCase() === myAsset.ticker);
      
      if (info && info.y > 0) {
          let estimatedYearly = myAsset.value * info.y;
          totalYearly += estimatedYearly;
          
          finalItems.push({
              ticker: myAsset.ticker,
              yieldPct: (info.y * 100).toFixed(2),
              nextMonth: info.m,
              estAmount: estimatedYearly
          });
      }
  });

  finalItems.sort((a,b) => b.estAmount - a.estAmount);

  return { totalYearly: totalYearly, items: finalItems };
}

/**
 * "Stock Battle" module: Compares two assets side-by-side.
 * Resolves ticker names (e.g., "Ferrari" -> "RACE") and scores them based on 
 * Valuation, Growth, Profitability, and Momentum.
 * * @param {string} inputA - Name or Ticker of the first asset.
 * @param {string} inputB - Name or Ticker of the second asset.
 * @return {Object} JSON comparison result including the winner, scores, and a verdict.
 */
function runStockBattle(inputA, inputB) {
  const API_KEY = GEMINI_API_KEY; 
  
  const MODELS = [
    "gemini-2.0-flash",      
    "gemini-1.5-flash",      
    "gemini-flash-latest"    
  ];
  
  // English Prompt
  const prompt = `
    You are a Senior Equity Analyst on Wall Street.
    
    USER INPUT:
    1. "${inputA}"
    2. "${inputB}"
    
    TASK 1 (IDENTIFICATION):
    Identify the companies and find their official TICKER (e.g., "Apple" -> "AAPL").
    
    TASK 2 (COMPARATIVE ANALYSIS):
    Perform a deep comparison based on current fundamentals:
    - Valuation (P/E, PEG, Price/Sales).
    - Profitability (Net Margins, FCF).
    - Growth (Revenue and Earnings Growth Y/Y).
    - Momentum/Risks.

    STRICT JSON OUTPUT (No Markdown):
    {
      "resolved_ticker_a": "TICKER_A",
      "resolved_ticker_b": "TICKER_B",
      "winner": "WINNING_TICKER",
      "scoreA": 75,
      "scoreB": 60,
      "strengths_a": ["Strength 1", "Strength 2", "Strength 3"],
      "strengths_b": ["Strength 1", "Strength 2", "Strength 3"],
      "verdict": "Detailed discursive analysis (approx 70-80 words). Technically explain why X wins over Y today. Cite key metrics."
    }
  `;

  let result = { 
    resolved_ticker_a: inputA, resolved_ticker_b: inputB, 
    winner: "N/A", scoreA: 50, scoreB: 50, 
    strengths_a: [], strengths_b: [],
    verdict: "Analysis unavailable." 
  };

  for (let m=0; m<MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
      
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        let txt = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        txt = txt.replace(/```json/g, "").replace(/```/g, "").trim();
        
        const firstBrace = txt.indexOf('{');
        const lastBrace = txt.lastIndexOf('}');
        if (firstBrace !== -1 && lastBrace !== -1) {
            txt = txt.substring(firstBrace, lastBrace + 1);
            result = JSON.parse(txt);
            break; 
        }
      }
    } catch(e) { console.warn(`Error with ${MODELS[m]}: ${e}`); }
  }

  return result;
}

/**
 * Advanced Asset Analysis Module (Investor AI).
 * Automatically detects if the input is a Stock or Crypto and applies the specific
 * institutional-grade framework provided by the user.
 * * @param {string} inputName - The name or ticker of the asset (e.g., "NVDA", "Ethereum").
 * @return {string} The formatted HTML/Markdown analysis.
 */
function analyzeAsset(inputName) {
  const API_KEY = GEMINI_API_KEY; // Assicurati che questa variabile globale sia accessibile
  
  // Model Sequence (Priority to 2.0 Flash for speed/quality balance)
  const MODELS = [
    "gemini-2.0-flash", 
    "gemini-1.5-flash", 
    "gemini-flash-latest"
  ];

  // --- 1. SYSTEM PROMPTS (Come da tua richiesta) ---

  const INVESTOR_PROMPT_STOCK = `
ROLE: You are an Elite Global Macro Strategist & Senior Equity Research Analyst. 
You combine the fundamental depth of Warren Buffett with the risk management of a Hedge Fund Manager.

OBJECTIVE: Provide a professional investment analysis for the company: "${inputName}".
OUTPUT LANGUAGE: Italian (Strictly).
TONE: Professional, Direct, Educational, Data-Driven.

*** FORMATTING & UX RULES (CRITICAL) ***
1. **Use Markdown:** Use **Bold** for key numbers and headers. Use Tables for data comparison.
2. **Educational Overlay:** Whenever you mention a complex metric (e.g., ROIC, Z-Score, SBC), you MUST provide a micro-explanation in parentheses explaining WHY it matters.
   - *Example:* "ROIC: 15% (Creating Value: Returns exceed cost of capital)."
   - *Example:* "Altman Z-Score: 1.2 (Distress Zone: High bankruptcy risk)."
3. **Structure:** Break the text into short paragraphs and bullet points. No walls of text.

--- ANALYSIS FRAMEWORK (CHAIN OF THOUGHT) ---

PHASE 1: 🚨 EXECUTIVE SUMMARY & CONTEXT
- **Thesis:** The "Elevator Pitch" in 2 sentences.
- **Macro Overlay:** Briefly mention if the current macro environment (Rates, Inflation) helps or hurts this specific stock.

PHASE 2: 🏥 FINANCIAL HEALTH & FORENSIC SCORECARD
*Create a Markdown Table with these columns: Metric | Value | Rating (Good/Bad) | Context/Explanation*
- **Piotroski F-Score (0-9):** Operational Efficiency.
- **Altman Z-Score:** Bankruptcy Risk.
- **Beneish M-Score:** Earnings Manipulation Check.
- **SBC % of Revenue:** Stock-Based Compensation (Dilution risk).
- **ROIC vs WACC:** Is the company creating value (ROIC > WACC) or destroying it?

PHASE 3: 🕵️ DEEP DIVE (THE "SILENT KILLERS")
- **Quality of Earnings:** Compare GAAP Net Income vs Non-GAAP. Is the difference massive? Explain if it's a "red flag".
- **Concentration Risk:** Does >10% of revenue come from ONE client? (Single Point of Failure).
- **Moat Analysis:** Is the competitive advantage durable? (Network Effect, Switching Costs).

PHASE 4: 🔮 VALUATION & SCENARIOS (12-MONTH VIEW)
- **Variant Perception:** What does the Market think vs. What do WE think? Where is the "Alpha"?
- **Scenario Table:** Create a table with 3 rows:
  1. **🐻 BEAR Case (20% Prob):** Recession/Execution Failure -> Price Target?
  2. **⚖️ BASE Case (50% Prob):** Consensus -> Price Target?
  3. **🐂 BULL Case (30% Prob):** Blue Sky Execution -> Price Target?
  *Calculate the Probability Weighted Expected Return.*

PHASE 5: 🛡️ RISK MANAGEMENT & ACTION PLAN
- **The "Pre-Mortem" (Inversion):** Assume it's 2026 and the stock is down 60%. Write the "Autopsy": Why did it die?
- **Technical Check:** Where are the Support/Resistance levels? Is it overbought (RSI > 70)?
- **Final Verdict:**
  - **RATING:** [STRONG BUY / BUY / HOLD / SELL / AVOID]
  - **ACTION:** Entry Price Range & Stop Loss Level (Crucial).
  - **POSITION SIZING:** Suggest % allocation based on Kelly Criterion (assume medium risk tolerance).
  - **"Change My Mind" Triggers:** 3 objective signals that would force us to sell.
`;

  const INVESTOR_PROMPT_CRYPTO = `
ROLE: You are an Elite Crypto Researcher & Tokenomics Expert.
OBJECTIVE: Provide a deep-dive analysis for the crypto project: "${inputName}".
OUTPUT LANGUAGE: Italian (Strictly).

*** FORMATTING RULES ***
Use Markdown, Bold for keys, and keep it highly readable.

--- CRYPTO ANALYSIS FRAMEWORK ---

PHASE 1: 🪙 IDENTITY & FOUNDATIONS
- **What is it?** Technology, Layer (L1/L2), Consensus mechanism.
- **Real World Application:** Does it solve a real problem?
- **VC & Founders:** Who is behind it? (Doxxed? VC Backed? Community led?).
- **Cycles:** Calculate price targets based on: 4-Year Cycle, 320-Day Cycle, 80-Day Cycle.

PHASE 2: 🧪 FORENSIC CHECK (6 KEY QUESTIONS)
1. Utility: Does the token make sense or is it just governance/memecoin?
2. Value Creation: How does it redefine the space vs Banks/TradFi?
3. Disruption: Which industry is it disrupting?
4. Team Track Record: Have they built successful tech before?
5. Polish: Is the website/docs professional?
6. Whitepaper: Innovative or copy-paste?

PHASE 3: 🌊 ECOSYSTEM & NARRATIVE
- **Narrative Fit:** Is it part of a hot trend (Restaking, RWA, AI, DePIN) for the next 6-12 months?
- **Adoption Curve:** Are we early or is it priced in?
- **Correlation:** How does it move vs ETH/BTC?

PHASE 4: ⛓️ ON-CHAIN HEALTH & TOKENOMICS
- **TVL Trend:** Last 90 days direction.
- **Unlocks:** Are there vesting cliffs coming? (Inflation risk).
- **Revenue:** Is the protocol generating REAL yield?
- **Moat:** Can it be forked easily? (e.g. Uniswap vs Sushi).

PHASE 5: 🗳️ GOVERNANCE & EXIT STRATEGY
- **Governance:** Who holds the power (DAO vs Insiders)?
- **Community:** Sentiment on Twitter/Discord.
- **Exit Strategy:**
  - "Fully Valued" Price Target?
  - Portfolio Fit: Is it a core hold or a cycle trade?
  - Action Plan: Entry Zone & Profit Taking Levels.

BONUS: Staking/Yield opportunities for this specific token.
`;

  // --- 2. ROUTER PROMPT (Detect Type) ---
  // We ask Gemini to classify the input first to choose the right framework.
  const ROUTER_PROMPT = `
    Classify the financial asset "${inputName}".
    Return ONLY one word: "STOCK" or "CRYPTO".
    If unsure or it's a commodity/ETF, treat as "STOCK".
    If it's Bitcoin, Ethereum, Solana, or any token, treat as "CRYPTO".
  `;

  // --- 3. EXECUTION ---
  
  // Step A: Determine Type
  let assetType = "STOCK"; // Default
  try {
    const routerPayload = { contents: [{ parts: [{ text: ROUTER_PROMPT }] }] };
    const routerRes = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`,
      { method: "post", contentType: "application/json", payload: JSON.stringify(routerPayload), muteHttpExceptions: true }
    );
    if(routerRes.getResponseCode() === 200) {
      const typeText = JSON.parse(routerRes.getContentText()).candidates[0].content.parts[0].text.trim().toUpperCase();
      if(typeText.includes("CRYPTO")) assetType = "CRYPTO";
    }
  } catch(e) {
    console.warn("Router failed, defaulting to STOCK. Error: " + e);
  }

  // Step B: Select Prompt based on Type
  const finalPrompt = (assetType === "CRYPTO") ? INVESTOR_PROMPT_CRYPTO : INVESTOR_PROMPT_STOCK;

  // Step C: Generate Analysis (Loop through models for robustness)
  for (let m = 0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { 
        contents: [{ parts: [{ text: finalPrompt }] }],
        generationConfig: {
            temperature: 0.7, // Creative but analytical
            maxOutputTokens: 8000 // Long form content
        }
      };
      
      const response = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        
        // Convert Markdown to simple HTML for better display in your web app if needed, 
        // or return Markdown if your frontend handles it. 
        // Here we return raw Markdown/Text which is standard for LLM responses.
        return text; 
      }
    } catch (e) {
      console.error(`Analysis Model ${MODELS[m]} failed: ${e}`);
    }
  }

  return "⚠️ Error: AI models are currently overloaded. Please try again shortly.";
}