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
 * Performs a deep-dive "Chief Risk Officer" assessment.
 * UPGRADE: Now analyzes Sector Concentration, Geographic Exposure, and Specific Crash Scenarios.
 * Uses aggressive caching (12 hours).
 * * @return {Object} JSON object with detailed risk metrics and strategic hedging advice.
 */
function getPortfolioRiskAnalysis() {
  const CACHE_KEY = "GEMINI_RISK_DEEP_V2"; // Updated Key Version
  const cache = CacheService.getScriptCache();
  
  // Check cache first
  const cachedResult = cache.get(CACHE_KEY);
  if (cachedResult) return JSON.parse(cachedResult);

  // 1. Retrieve Data
  const dash = getDashboardData();
  const portfolio = getLivePortfolio();
  
  // 2. Pre-process Data for Context (Aggregation)
  // We aggregate sectors and countries to give the AI a "Macro View" of the portfolio
  let sectorMap = {};
  let countryMap = {};
  let totalEquity = 0;

  const allAssets = [...portfolio.stocks, ...portfolio.etfs];
  
  allAssets.forEach(a => {
    let val = parseFloat(String(a.val).replace(/[€$£%\s]/g, '').replace(',', '.')) || 0;
    if (val > 0) {
      totalEquity += val;
      let s = a.sector || "Other";
      let c = a.ctry || "Global";
      sectorMap[s] = (sectorMap[s] || 0) + val;
      countryMap[c] = (countryMap[c] || 0) + val;
    }
  });

  // Format Top 5 Sectors/Countries for the Prompt
  const formatMap = (map) => Object.entries(map)
    .sort((a,b) => b[1] - a[1])
    .slice(0, 5)
    .map(([k,v]) => `${k} (${((v/totalEquity)*100).toFixed(0)}%)`)
    .join(", ");

  const topSectors = formatMap(sectorMap);
  const topCountries = formatMap(countryMap);
  
  // Clean List for top individual positions
  const cleanList = (list) => list.slice(0, 10).map(i => `${i.t} (${i.pct})`).join(", ");
  const topStocks = cleanList(portfolio.stocks);
  const topCrypto = cleanList(portfolio.crypto);

  // 3. Construct the "Chief Risk Officer" Prompt
  const prompt = `
    ROLE: You are the Chief Risk Officer (CRO) of a Multi-Family Office.
    Your job is NOT to be nice. Your job is to protect capital.
    
    PORTFOLIO SNAPSHOT:
    - Liquid Cash Buffer: ${dash.summary.cash.percent} (Total: ${dash.liquidNetWorth})
    - Stock Exposure: ${dash.summary.stocks.percent}
    - Crypto Exposure: ${dash.summary.crypto.percent}
    - ETFs (Passive): ${dash.summary.etfs.percent}
    
    RISK FACTORS (Calculated):
    - Top Sectors Exposure: [${topSectors}]
    - Geographic Exposure: [${topCountries}]
    - Top Volatile Positions: ${topStocks}
    - Crypto Holdings: ${topCrypto}

    TASK: Perform a Stress Test & Asset Allocation Review.
    
    OUTPUT FORMAT: STRICT JSON (No Markdown).
    Language: Italian.

    JSON STRUCTURE:
    {
      "riskScore": "1-100",  // 1=Treasury Bills, 100=Degen Leverage
      "riskLevel": "Low/Medium/High/Extreme",
      "summary": "1 sharp sentence summarizing the main vulnerability.",
      "concentration": "Detailed analysis of Sector/Country bias. Are we too exposed to Tech or USA? Is diversification real or fake?",
      "stressTest": {
        "marketCrash": "Est. Portfolio Drawdown if S&P500 falls 20% (e.g., -25%)",
        "cryptoWinter": "Est. Portfolio Drawdown if BTC falls 50% (e.g., -10%)"
      },
      "suggestions": [
        "Actionable Advice 1 (e.g., 'Reduce NVDA by 5% to rebalance Tech')",
        "Actionable Advice 2 (e.g., 'Increase Gold/Bonds as hedge')",
        "Actionable Advice 3 (Strategic view)"
      ]
    }
    
    GUIDELINES:
    - If Cash is > 30%, warn about "Inflation Risk" (Opportunity Cost).
    - If Crypto is > 20%, warn about "Extreme Volatility".
    - If Top Sector (e.g. Tech) is > 40%, warn about "Concentration Risk".
    - Be numeric and specific in the suggestions.
  `;

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  
  let finalResult = { 
    riskLevel: "N/A", 
    riskScore: 0, 
    summary: "AI Analysis unavailable", 
    stressTest: { marketCrash: "--%", cryptoWinter: "--%" },
    concentration: "No data",
    suggestions: ["Retry later"]
  };

  for (let m = 0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const res = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        text = text.replace(/```json/g, "").replace(/```/g, "").trim();
        
        // Find JSON boundaries to avoid parsing errors
        const s = text.indexOf('{');
        const e = text.lastIndexOf('}');
        if (s !== -1 && e !== -1) {
            finalResult = JSON.parse(text.substring(s, e + 1));
            
            // Save to cache for 12 hours (High-value, low-frequency report)
            cache.put(CACHE_KEY, JSON.stringify(finalResult), 43200); 
            break; 
        }
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
 * UPDATED: Handles both fast polling (Cache) and forced user refresh (Live).
 *
 * @param {boolean} onlyMacro - If true (Polling), uses cache. If false (User Click), FORCES new analysis.
 */
function getMarketInsightsData(onlyMacro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = CacheService.getScriptCache();
  
  const MACRO_CACHE_KEY = "MARKET_MACRO_DATA_V1";
  const AI_CACHE_KEY = "MARKET_AI_INSIGHTS_PERSIST_V1"; 
  
  // --- 1. MACRO DATA (Prices & Indices) ---
  let macro = null;
  const cachedJSON = cache.get(MACRO_CACHE_KEY);

  // LOGIC CHANGE: If user clicks Refresh (!onlyMacro), we ignore cache and fetch fresh prices from Sheet
  if (cachedJSON && onlyMacro) {
      macro = JSON.parse(cachedJSON);
  } else {
      macro = { spx: 0, dow: 0, nasdaq: 0, russell: 0, vix: 0, us10y: 0, cryptoTrend: "N/A", me: 0 };
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
          // Save for 60s
          cache.put(MACRO_CACHE_KEY, JSON.stringify(macro), 60);
        }
      } catch(e) { console.error("Error Dashboard Data: " + e); }
  }

  // --- 2. AI MEMORY (Cache) ---
  let cachedAiRaw = cache.get(AI_CACHE_KEY);
  let aiData = cachedAiRaw ? JSON.parse(cachedAiRaw) : {
      sentiment: { score: 5, label: "Neutral" },
      analysis: { macro: "Click refresh to analyze.", portfolio: "..." },
      market_events: [],
      portfolio_events: []
  };

  // IF POLLING (onlyMacro = true): Return cached AI immediately (don't run Gemini)
  if (onlyMacro) {
      return {
          ...macro,
          metrics: { beta: 0, pe: 0 }, 
          sentiment: aiData.sentiment,
          analysis: aiData.analysis,
          market_events: aiData.market_events,
          portfolio_events: aiData.portfolio_events
      };
  }

  // =========================================================
  // USER CLICKED REFRESH -> RUNNING FULL ANALYSIS (NO CACHE)
  // =========================================================

  // --- 3. PORTFOLIO METRICS ---
  const port = getLivePortfolio(); 
  
  let cryptoTotalChange = 0; let cryptoCount = 0;
  if (port.crypto && port.crypto.length > 0) {
    port.crypto.forEach(c => {
      if (c.pct) { 
        let val = parseFloat(String(c.pct).replace(/[€$£%\s]/g, '').trim().replace(',', '.')) || 0;
        cryptoTotalChange += val; cryptoCount++;
      }
    });
    macro.cryptoTrend = (cryptoCount > 0 ? (cryptoTotalChange / cryptoCount).toFixed(2) : "0") + "%";
  }

  const equityAssets = [...port.stocks, ...port.etfs];
  let totalEquityVal = 0; let weightedBeta = 0; let weightedPE = 0; let peEligibleVal = 0;
  let sectorMap = {}; let countryMap = {};

  const parseVal = (v) => parseFloat(String(v).replace(/[^\d.-]/g, '').replace(',', '.')) || 0;

  const enrichedEquity = equityAssets.map(a => {
    const val = parseVal(a.val);
    const beta = parseFloat(a.beta) || 1; 
    const pe = parseFloat(a.pe) || 0;
    totalEquityVal += val;
    weightedBeta += (beta * val);
    if (pe > 0 && a.sector !== 'ETFs') { weightedPE += (pe * val); peEligibleVal += val; }
    const s = a.sector || "Other"; sectorMap[s] = (sectorMap[s] || 0) + val;
    const c = a.ctry || "Global"; countryMap[c] = (countryMap[c] || 0) + val;
    return { t: a.t, val: val, dCh: a.dCh, details: { beta: beta } };
  }).sort((a, b) => b.val - a.val); 

  const portBeta = totalEquityVal > 0 ? (weightedBeta / totalEquityVal).toFixed(2) : 1;
  const portPE = peEligibleVal > 0 ? (weightedPE / peEligibleVal).toFixed(1) : "N/A";

  // --- 4. AI GENERATION (Gemini) ---
  let aiRes = { 
    sentiment_score: 5, sentiment_label: "Neutral", 
    macro_analysis: null, portfolio_analysis: null, 
    market_events: [], portfolio_events: [] 
  };

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  const today = new Date().toISOString().split('T')[0];
  
  const topSectors = Object.entries(sectorMap).sort((a,b) => b[1]-a[1]).slice(0,5).map(([k,v]) => k).join(", ");
  const topCountries = Object.entries(countryMap).sort((a,b) => b[1]-a[1]).slice(0,3).map(([k,v]) => k).join(", ");
  const myTickers = enrichedEquity.map(a => a.t).join(", ");
  const assetsStr = enrichedEquity.slice(0, 30).map(a => `${a.t} (${a.dCh})`).join("\n");

  // PROMPT "DEEP DIVE" (10 DAYS)
  const prompt = `
    Role: Institutional Hedge Fund Manager & Senior Macro Strategist.
    Date: ${today}.
    
    CONTEXT DATA:
    MACRO: S&P500 ${macro.spx}% | VIX ${macro.vix} | US10Y ${macro.us10y}%.
    PORTFOLIO: Beta ${portBeta} | P/E ${portPE} | Top Sectors: ${topSectors} | Top Countries: ${topCountries}.
    ASSETS: \n${assetsStr}
    
    MISSION:
    Perform a "Deep Dive" risk & opportunity analysis. Be extremely specific, critical, and forward-looking.
    
    TASK LIST:
    
    A. MACRO SYNTHESIS (The "Why"):
    Analyze how Rates, Inflation, and Geopolitics are impacting the market RIGHT NOW. Connect dots.
    
    B. PORTFOLIO DIAGNOSIS:
    Ruthless review of my allocation. Am I too exposed to Tech? Too defensive? What is my biggest blind spot?
    
    C. SENTIMENT SCORE (0-10):
    Based on VIX, and Macro. (0=Extreme Fear, 10=Extreme Greed).
    
    D. EVENTS & CATALYSTS (Next 10 Days - STRICT & COMPREHENSIVE):
    I need EVERY significant event that could move my money in the next 10-15 days.
    
    1. MACRO EVENTS: List ALL critical economic data (CPI, PPI, Jobs, GDP), Central Bank meetings (Fed, ECB), or Geopolitical deadlines.
    2. PORTFOLIO EVENTS: List ALL Earnings, Dividends, Product Launches, or Governance votes specifically for [${myTickers}].
    
    *** CRITICAL INSTRUCTION FOR "event" FIELD ***
    Do NOT just list the name. You MUST provide the "Analytic Context" and "Impact Prediction".
    - BAD: "US CPI Release"
    - GOOD: "US CPI: Critical for Fed Pivot. If >3.2%, expect Tech sell-off. High Volatility."
    - GOOD: "NVDA Earnings: Focus on Data Center guidance. +/- 8% implied move."

    JSON OUTPUT FORMAT:
    {
      "sentiment_score": 5.5,
      "sentiment_label": "Fear/Neutral/Greed",
      "macro_analysis": "Deep institutional commentary here...",
      "portfolio_analysis": "Specific, actionable portfolio advice here...",
      "market_events": [{"ticker": "MACRO", "event": "Event Name: Context & Impact Prediction", "date": "YYYY-MM-DD"}],
      "portfolio_events": [{"ticker": "TICKER", "event": "Event Name: Context & Impact Prediction", "date": "YYYY-MM-DD"}]
    }
  `;

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

  // SAVE NEW ANALYSIS TO CACHE
  const finalResult = {
    sentiment: { score: aiRes.sentiment_score, label: aiRes.sentiment_label },
    analysis: { macro: aiRes.macro_analysis, portfolio: aiRes.portfolio_analysis },
    market_events: aiRes.market_events,
    portfolio_events: aiRes.portfolio_events
  };

  if (aiRes.market_events.length > 0 || aiRes.portfolio_events.length > 0) {
      cache.put(AI_CACHE_KEY, JSON.stringify(finalResult), 21600); // 6 Hours
  }

  return {
    ...macro,
    metrics: { beta: portBeta, pe: portPE },
    ...finalResult
  };
}

/**
 * Performs a deep-dive "Chief Risk Officer" assessment.
 * UPGRADE: Enhanced Analytical Prompt for Asset Allocation & Correlation.
 * Checks for "Fake Diversification" and specific hedging strategies.
 *
 * @param {boolean} forceRefresh - If true, bypasses cache and forces new AI analysis.
 */
function getPortfolioRiskAnalysis(forceRefresh) {
  const CACHE_KEY = "GEMINI_RISK_DEEP_V2"; 
  const cache = CacheService.getScriptCache();
  
  // 1. CACHE CHECK (Skipped if forceRefresh is true)
  if (!forceRefresh) {
      const cachedResult = cache.get(CACHE_KEY);
      if (cachedResult) return JSON.parse(cachedResult);
  }

  // 2. DATA AGGREGATION
  const dash = getDashboardData();
  const portfolio = getLivePortfolio();
  
  let sectorMap = {};
  let countryMap = {};
  let totalEquity = 0;

  const allAssets = [...portfolio.stocks, ...portfolio.etfs];
  
  allAssets.forEach(a => {
    let val = parseFloat(String(a.val).replace(/[€$£%\s]/g, '').replace(',', '.')) || 0;
    if (val > 0) {
      totalEquity += val;
      let s = a.sector || "Other";
      let c = a.ctry || "Global";
      sectorMap[s] = (sectorMap[s] || 0) + val;
      countryMap[c] = (countryMap[c] || 0) + val;
    }
  });

  const formatMap = (map) => Object.entries(map)
    .sort((a,b) => b[1] - a[1])
    .slice(0, 5)
    .map(([k,v]) => `${k} (${((v/totalEquity)*100).toFixed(0)}%)`)
    .join(", ");

  const topSectors = formatMap(sectorMap);
  const topCountries = formatMap(countryMap);
  const cleanList = (list) => list.slice(0, 10).map(i => `${i.t} (${i.pct})`).join(", ");
  const topStocks = cleanList(portfolio.stocks);
  const topCrypto = cleanList(portfolio.crypto);

  // 3. THE "ELITE CRO" PROMPT
  const prompt = `
    ROLE: You are the Chief Risk Officer (CRO) of a Top-Tier Multi-Family Office.
    Your methodology is based on Ray Dalio's "All Weather" principles and Taleb's Risk Management.
    
    PORTFOLIO STRUCTURE:
    - 💵 CASH / LIQUIDITY: ${dash.summary.cash.percent} (Value: ${dash.liquidNetWorth})
    - 📈 EQUITY (Stocks): ${dash.summary.stocks.percent}
    - 📉 ETFS (Passive): ${dash.summary.etfs.percent}
    - ⚡ CRYPTO (High Vol): ${dash.summary.crypto.percent}
    
    DEEP EXPOSURE DATA (Equity Component):
    - Sector Dominance: [${topSectors}]
    - Geographic Bias: [${topCountries}]
    - Top Positions: ${topStocks}
    - Crypto Holdings: ${topCrypto}

    MISSION:
    Conduct a forensic analysis of the Asset Allocation quality. Don't just read the numbers; interpret the CORRELATIONS.
    
    KEY ANALYTICAL TASKS:
    
    1. "FAKE DIVERSIFICATION" CHECK:
       - Do I own different names that act the same? (e.g. Tech Stocks + Nasdaq ETF + Crypto = 100% Correlation).
       - Identify the "Single Point of Failure" (The one factor that kills the portfolio).
    
    2. EFFICIENCY & SIZING:
       - Is the Cash drag too high given inflation?
       - Is the Crypto allocation reckless (>5-10%) or strategic?
       - Is there Home Country Bias?

    3. STRESS TEST SIMULATION:
       - Calculate expected drawdown based on the weight of High-Beta assets (Crypto/Tech) vs Low-Beta (Cash/Bonds).

    OUTPUT FORMAT: STRICT JSON (No Markdown).

    JSON STRUCTURE:
    {
      "riskScore": "1-100", // (1=Safe, 100=Reckless)
      "riskLevel": "Low/Medium/High/Extreme",
      "summary": "1 brutal sentence on the portfolio's main weakness.",
      "concentration": "Detailed analysis. Discuss Correlation, Sector Overlap, and 'Fake Diversification'. Be specific about which assets are overlapping.",
      "stressTest": {
        "marketCrash": "Est. Portfolio Drawdown if S&P500 falls 20% (e.g. -12%). Explain logic briefly.",
        "cryptoWinter": "Est. Portfolio Drawdown if Bitcoin falls 50% (e.g. -5%)."
      },
      "suggestions": [
        "Rebalancing Action 1 (Specific % move, e.g. 'Cut Tech by 10%')",
        "Hedging Strategy (e.g. 'Buy Gold/Bonds to de-correlate')",
        "Optimization (e.g. 'Deploy Cash into Dividend Aristocrats')"
      ]
    }
  `;

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  
  let finalResult = { 
    riskLevel: "N/A", riskScore: 0, summary: "AI unavailable", 
    stressTest: { marketCrash: "--%", cryptoWinter: "--%" },
    concentration: "No data", suggestions: ["Retry later"]
  };

  for (let m = 0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const res = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        text = text.replace(/```json/g, "").replace(/```/g, "").trim();
        
        const s = text.indexOf('{');
        const e = text.lastIndexOf('}');
        if (s !== -1 && e !== -1) {
            finalResult = JSON.parse(text.substring(s, e + 1));
            // Save to cache for 12 hours
            cache.put(CACHE_KEY, JSON.stringify(finalResult), 43200); 
            break; 
        }
      }
    } catch (e) { console.error("Risk AI Error: " + e.toString()); }
  }

  return finalResult;
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
    Role: Algorithmic Hedge Fund Manager. Date: ${today}.
    MACRO: S&P500 ${macro.spx}% | VIX ${macro.vix} | US10Y ${macro.us10y}% | Crypto Trend ${macro.cryptoTrend}.
    PORTFOLIO: Beta ${portBeta} | P/E ${portPE} | Top Sectors: ${topSectors} | Top Countries: ${topCountries}.
    ASSETS: \n${assetsStr}
    
    TASK:
    A. Analyze Macro interaction (Rates, Inflation, Geopolitics).
    B. Diagnose Portfolio Style & Risks.
    C. Sentiment Score (0-10) based on VIX and Momentum.
    
    D. DEEP DIVE EVENTS & CATALYSTS (Next 30 Days):
    Identify 5 CRITICAL Global Market Events (Fed, ECB, CPI, Jobs) + Specific Earnings/Divs for my assets: [${myTickers}].
    
    CRITICAL INSTRUCTION FOR "event" FIELD:
    Do NOT just list the name. You MUST include the "Why it matters" and "Expected Impact".
    - BAD: "Apple Earnings"
    - GOOD: "Apple Earnings: Focus on China sales decline. If miss, expect -5% drop. High Volatility."
    - BAD: "US CPI"
    - GOOD: "US CPI Data: Crucial for Fed Pivot. If >3.2%, Tech stocks will suffer. Watch 10Y Yield."

    JSON OUTPUT:
    {
      "sentiment_score": 5.5,
      "sentiment_label": "Fear/Neutral/Greed",
      "macro_analysis": "Detailed view on how macro factors (Rates/War/Oil) are impacting the market NOW. Be specific.",
      "portfolio_analysis": "Specific feedback on my allocation. Am I too exposed to Tech? Too defensive? Give actionable advice.",
      "market_events": [{"ticker": "MACRO", "event": "Event Name: Analytic Context & Impact Prediction", "date": "YYYY-MM-DD"}],
      "portfolio_events": [{"ticker": "TICKER", "event": "Event Name: Analytic Context & Impact Prediction", "date": "YYYY-MM-DD"}]
    }
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
 * NOW WITH JSON-BASED SMART TICKER RESOLUTION.
 * Fixes "Netflix" issues by forcing strict JSON output for ticker identification.
 * * @param {string} inputName - The name or ticker (e.g., "Netflix" or "NFLX").
 * @param {number|string} [currentPrice] - Optional real-time price.
 * @return {string} The formatted HTML/Markdown analysis.
 */
function analyzeAsset(inputName, currentPrice) {
  const API_KEY = GEMINI_API_KEY; 
  
  // --- 1. CONTEXT: DATE & TIME ---
  const today = new Date().toLocaleDateString('it-IT', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  
  // --- 2. SMART TICKER & PRICE RESOLUTION ---
  let resolvedTicker = inputName;
  
  if (!currentPrice) {
    // A. Try fetching with input as is
    currentPrice = fetchPriceYahoo(inputName);

    // B. If failed, ask AI to find the Ticker using STRICT JSON
    if (!currentPrice) {
       console.log(`Price miss for '${inputName}'. Resolving ticker via AI...`);
       try {
         const tickerPrompt = `
           Identify the financial ticker for "${inputName}".
           Return a STRICT JSON object: {"symbol": "THE_TICKER"}.
           If it is a crypto, append "-USD" (e.g. "BTC-USD").
           Example: {"symbol": "NFLX"}
           ONLY JSON. NO TEXT.
         `;
         
         const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`;
         const payload = { contents: [{ parts: [{ text: tickerPrompt }] }] };
         const response = UrlFetchApp.fetch(url, {
           method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
         });
         
         if (response.getResponseCode() === 200) {
            let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
            
            // --- FIX ROBUSTEZZA JSON ---
            // Cerca la prima parentesi graffa aperta e l'ultima chiusa
            const jsonMatch = text.match(/\{[\s\S]*\}/);
            
            if (jsonMatch) {
                const json = JSON.parse(jsonMatch[0]); // Parla solo la parte JSON
                const aiTicker = json.symbol;
                
                console.log(`✅ AI Resolved '${inputName}' to '${aiTicker}'`);
                
                // C. Retry fetch with the clean resolved ticker
                const priceCheck = fetchPriceYahoo(aiTicker);
                
                // Aggiorna sempre il ticker risolto
                resolvedTicker = aiTicker;

                if (priceCheck) {
                  currentPrice = priceCheck; 
                }
            } else {
                console.warn("AI response did not contain valid JSON: " + text);
            }
            // ---------------------------
         }
       } catch(e) {
         console.warn("Ticker resolution failed: " + e);
       }
    }
  }

  // --- 3. PRICE ANCHORING CONTEXT ---
  // Header injection to verify data source visibly
  const statusHeader = currentPrice 
          ? `✅ **DATI DI MERCATO VERIFICATI**\n> **Asset:** ${resolvedTicker}\n> **Prezzo:** $${currentPrice}\n> **Data:** ${today}\n\n---\n\n`
          : `⚠️ **DATI DI MERCATO NON DISPONIBILI**\n> Prezzo non trovato per "${resolvedTicker}". L'analisi si basa su stime.\n\n---\n\n`;

  const priceContext = currentPrice 
    ? `REAL-TIME MARKET DATA (Verified): Price for ${resolvedTicker} is ${currentPrice}. USE THIS PRICE as t0 for all valuation models.` 
    : "REAL-TIME PRICE: UNAVAILABLE. You MUST estimate valuation based on the LAST KNOWN CLOSING PRICE you assume, but explicitly flag it as an estimate.";

  // Model Sequence
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];

  // --- 4. SYSTEM PROMPTS (ORIGINAL TEXT) ---

  const INVESTOR_PROMPT_STOCK = `
ROLE: You are an Elite Global Macro Strategist & Senior Equity Research Analyst. 
You combine the fundamental depth of Warren Buffett with the risk management of a Hedge Fund Manager.

DATE OF ANALYSIS: ${today}. (If your internal clock says 2024, IGNORE IT. Today is ${today}).

OBJECTIVE: Provide a professional investment analysis for the company: "${resolvedTicker}" (User searched: "${inputName}").
${priceContext}
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
- **Reference Price:** ${currentPrice ? currentPrice : "N/A (See Real-Time Data)"}
- **Variant Perception:** What does the Market think vs. What do WE think? Where is the "Alpha"?
- **Scenario Table:** Create a table with 3 rows:
  1. **🐻 BEAR Case (20% Prob):** Recession/Execution Failure -> Price Target?
  2. **⚖️ BASE Case (50% Prob):** Consensus -> Price Target?
  3. **🐂 BULL Case (30% Prob):** Blue Sky Execution -> Price Target?
  *Calculate the Probability Weighted Expected Return.*

PHASE 5: 🛡️ RISK MANAGEMENT & ACTION PLAN
- **The "Pre-Mortem" (Inversion):** Assume it's a future date and the stock is down 60%. Write the "Autopsy": Why did it die?
- **Technical Check:** Where are the Support/Resistance levels? Is it overbought (RSI > 70)?
- **Final Verdict:**
  - **RATING:** [STRONG BUY / BUY / HOLD / SELL / AVOID]
  - **ACTION:** Entry Price Range & Stop Loss Level (Crucial).
  - **POSITION SIZING:** Suggest % allocation based on Kelly Criterion (assume medium risk tolerance).
  - **"Change My Mind" Triggers:** 3 objective signals that would force us to sell.
`;

  const INVESTOR_PROMPT_CRYPTO = `
ROLE: You are an Elite Crypto Researcher & Tokenomics Expert.

DATE OF ANALYSIS: ${today}.
OBJECTIVE: Provide a deep-dive analysis for the crypto project: "${resolvedTicker}".
${priceContext}
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
  - "Fully Valued" Price Target? (Reference: ${currentPrice || 'Current Price'})
  - Portfolio Fit: Is it a core hold or a cycle trade?
  - Action Plan: Entry Zone & Profit Taking Levels.

BONUS: Staking/Yield opportunities for this specific token.
`;

  // --- ROUTER (Classify Asset) ---
  const ROUTER_PROMPT = `
    Classify the financial asset "${resolvedTicker}".
    Return ONLY one word: "STOCK" or "CRYPTO".
    If unsure or it's a commodity/ETF, treat as "STOCK".
    If it's Bitcoin, Ethereum, Solana, or any token, treat as "CRYPTO".
  `;
  
  let assetType = "STOCK";
  try {
    const routerPayload = { contents: [{ parts: [{ text: ROUTER_PROMPT }] }] };
    const routerRes = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`,
      { method: "post", contentType: "application/json", payload: JSON.stringify(routerPayload), muteHttpExceptions: true }
    );
    if(routerRes.getResponseCode() === 200) {
      const txt = JSON.parse(routerRes.getContentText()).candidates[0].content.parts[0].text.trim().toUpperCase();
      if(txt.includes("CRYPTO")) assetType = "CRYPTO";
    }
  } catch(e) {}

  const finalPrompt = (assetType === "CRYPTO") ? INVESTOR_PROMPT_CRYPTO : INVESTOR_PROMPT_STOCK;

  // --- GENERATION LOOP ---
  for (let m = 0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { 
        contents: [{ parts: [{ text: finalPrompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 8000 }
      };
      
      const response = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        return statusHeader + text; 
      }
    } catch (e) { console.error(`Model ${MODELS[m]} failed: ${e}`); }
  }

  return "⚠️ Error: AI models overloaded.";
}

/**
 * MASTER PRICE FETCHER
 * Strategy:
 * 1. Try Yahoo Finance v7 (Fastest).
 * 2. If blocked (401/403) or fails, fallback to GOOGLEFINANCE (Rock solid for Stocks).
 */
function fetchPriceYahoo(ticker) {
  let price = null;

  // --- ATTEMPT 1: YAHOO FINANCE (v7 Quote) ---
  try {
    const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${ticker}`;
    const params = {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36" }
    };
    const response = UrlFetchApp.fetch(url, params);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.quoteResponse && json.quoteResponse.result && json.quoteResponse.result.length > 0) {
        const data = json.quoteResponse.result[0];
        price = data.regularMarketPrice || data.postMarketPrice || data.preMarketPrice;
        console.log(`✅ Price found via Yahoo: ${price}`);
        return price;
      }
    } else {
      console.warn(`Yahoo Blocked (${code}) for ${ticker}. Switching to fallback...`);
    }
  } catch (e) {
    console.warn("Yahoo Fetch Crash: " + e);
  }

  // --- ATTEMPT 2: GOOGLE FINANCE FALLBACK (The "Sheet Bridge") ---
  // Only triggers if Yahoo fails. 100% success rate for Stocks/ETFs.
  if (!price) {
     console.log(`🔄 Yahoo failed. Attempting Google Finance fallback for '${ticker}'...`);
     price = fetchPriceGoogle(ticker);
  }

  return price;
}

/**
 * Fallback function that uses the actual Spreadsheet to calculate the price.
 * Uses the 'Config' sheet to perform a temporary calculation.
 */
function fetchPriceGoogle(ticker) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Use the "Config" sheet (which surely exists in your setup)
    let sheet = ss.getSheetByName("Config"); 
    if (!sheet) {
      // If missing, use the first available sheet
      sheet = ss.getSheets()[0]; 
    }

    // Use a distant, safe cell (e.g., Z100) to avoid overwriting data
    const cell = sheet.getRange("Z100"); 
    
    // 1. Write the native Google Sheets formula
    cell.setFormula(`=GOOGLEFINANCE("${ticker}"; "price")`);
    
    // 2. Force immediate sheet update
    SpreadsheetApp.flush();
    
    // 3. Read the calculated value
    const val = cell.getValue();
    
    // 4. Clear the cell (leave no traces)
    cell.clearContent();
    SpreadsheetApp.flush(); // Commit the cleanup

    // Check if it is a valid number
    if (typeof val === 'number' && !isNaN(val)) {
      console.log(`✅ Price found via GoogleFinance Bridge: ${val}`);
      return val;
    } else {
      console.warn(`GoogleFinance returned invalid data: ${val} (Is ticker correct?)`);
    }

  } catch (e) {
    console.warn("Google Finance Fallback Failed: " + e);
  }
  return null;
}