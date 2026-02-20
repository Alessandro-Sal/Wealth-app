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
  
  Expected Output: Language Italian and strictly valid JSON (no markdown, no backticks):
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
 * Solved P/E dependency: AI now infers Growth/Value and Cap Size from Tickers.
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

  // If user clicks Refresh (!onlyMacro), ignore cache and fetch fresh prices
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

  // IF POLLING (onlyMacro = true): Return cached AI immediately
  if (onlyMacro) {
      return {
          ...macro,
          metrics: { beta: 0, pe: 0 }, // Beta & PE are recalculated on fresh analysis
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
    
    // We still calculate PE if available for purely internal metrics, but AI won't strictly rely on it.
    if (pe > 0 && a.sector !== 'ETFs') { weightedPE += (pe * val); peEligibleVal += val; }
    
    const s = a.sector || "Other"; sectorMap[s] = (sectorMap[s] || 0) + val;
    const c = a.ctry || "Global"; countryMap[c] = (countryMap[c] || 0) + val;
    return { t: a.t, val: val, dCh: a.dCh, details: { beta: beta } };
  }).sort((a, b) => b.val - a.val); 

  const portBeta = totalEquityVal > 0 ? (weightedBeta / totalEquityVal).toFixed(2) : 1;
  const portPE = peEligibleVal > 0 ? (weightedPE / peEligibleVal).toFixed(1) : "N/A";

  // --- 4. AI GENERATION (Gemini) ---
  const today = new Date().toISOString().split('T')[0];
  const topSectors = Object.entries(sectorMap).sort((a,b) => b[1]-a[1]).slice(0,5).map(([k,v]) => k).join(", ");
  const topCountries = Object.entries(countryMap).sort((a,b) => b[1]-a[1]).slice(0,3).map(([k,v]) => k).join(", ");
  const myTickers = enrichedEquity.map(a => a.t).join(", ");
  const assetsStr = enrichedEquity.slice(0, 30).map(a => `${a.t} (${a.dCh})`).join("\n");

  const prompt = `
<role>
You are an Institutional Hedge Fund Manager & Senior Macro Strategist. Today is ${today}.
You are blunt, highly analytical, and focus on forward-looking catalysts and risk asymmetry.
</role>

<context_data>
[MACRO ENVIRONMENT]
- S&P500: ${macro.spx}%
- VIX: ${macro.vix}
- US10Y: ${macro.us10y}%

[PORTFOLIO SNAPSHOT]
- Portfolio Beta: ${portBeta}
- Partial P/E (Incomplete data): ${portPE}
- Top Sectors: ${topSectors}
- Top Countries: ${topCountries}

[TOP 30 ACTIVE HOLDINGS (Ticker & Daily Change)]
${assetsStr}
</context_data>

<mission>
Perform a "Deep Dive" risk & opportunity analysis based on the current context. Be specific, critical, and explicitly connect macro variables to the specific portfolio holdings.
</mission>

<task_list>
A. MACRO SYNTHESIS (The "Why"): Analyze how Rates, Inflation, and Geopolitics are impacting the market right now.
B. PORTFOLIO DIAGNOSIS: Ruthlessly review the allocation. Since P/E data is incomplete, deduce the Style (Growth vs Value) and Cap Size (Mega-Cap vs Small-Cap) strictly from the provided tickers. Are we too exposed to Mega-Cap Tech? Too defensive given the Beta? Identify the biggest blind spot.
C. EVENTS & CATALYSTS: Provide EVERY significant event (Macro and Portfolio-specific) that could move this money in the next 10-15 days.
</task_list>

<guidelines>
- The Sentiment Score must be between 0 (Extreme Fear) and 10 (Extreme Greed) based strictly on VIX and macro movements.
- For events, DO NOT just list the name. Provide "Analytic Context" and "Impact Prediction" (e.g., "US CPI: Critical for Fed Pivot. If >3.2%, expect Tech sell-off").
- "market_events" should focus on CPI, PPI, Central Banks, or broad geopolitical deadlines.
- "portfolio_events" should focus specifically on earnings, dividends, or product launches for [${myTickers}].
- Output Language: Keep JSON keys in English. Write all string values (analysis, labels, event descriptions) in Italian.
</guidelines>

<output_format>
{
  "sentiment": {
    "score": "Float (0.0 to 10.0)",
    "label": "Fear | Neutral | Greed"
  },
  "analysis": {
    "macro": "Deep institutional commentary on the global environment...",
    "portfolio": "Specific, actionable diagnosis based on Beta, Sector mix, and implied Style/Cap Size of the given tickers..."
  },
  "market_events": [
    {"ticker": "MACRO", "event": "Name: Context & Impact Prediction", "date": "YYYY-MM-DD"}
  ],
  "portfolio_events": [
    {"ticker": "TICKER", "event": "Name: Context & Impact Prediction", "date": "YYYY-MM-DD"}
  ]
}
</output_format>
  `;

  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  
  // Default fallback if generation fails
  let finalResult = {
      sentiment: { score: 5, label: "Neutral" },
      analysis: { macro: "Analisi Macro non disponibile. Riprova più tardi.", portfolio: "Analisi Portafoglio non disponibile." },
      market_events: [],
      portfolio_events: []
  };

  for (let m=0; m<MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      
      const payload = { 
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          responseMimeType: "application/json"
        }
      };
      
      const options = { 
        method: "post", 
        contentType: "application/json", 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      };
      
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        let txt = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
        txt = txt.replace(/```json/g, "").replace(/```/g, "").trim();
        finalResult = JSON.parse(txt);
        
        // Save to cache on successful parse
        if (finalResult.market_events && finalResult.market_events.length > 0) {
          cache.put(AI_CACHE_KEY, JSON.stringify(finalResult), 21600); // 6 Hours
        }
        break;
      }
    } catch(err) { console.warn("AI Err: " + err); }
  }

  return {
    ...macro,
    metrics: { beta: portBeta, pe: portPE },
    ...finalResult
  };
}

/**
 * Automates the nightly email dispatch of the Market Insights report.
 * Checks for weekdays (Mon-Fri) to ensure it only runs on trading days.
 */
function sendNightlyMarketReport() {
  const recipient = "alessandro.saladino01@gmail.com";
  const today = new Date();
  const day = today.getDay(); // 0 = Sunday, 6 = Saturday

  // Execute only on Weekdays (Monday=1 to Friday=5)
  if (day === 0 || day === 6) {
    console.log("Weekend: Skipping nightly report.");
    return;
  }

  try {
    // Force fresh analysis (false) to get the latest data
    const insights = getMarketInsightsData(false);

    // Prepare HTML Email Body
    // Using safe access to properties to prevent crashes if AI fails
    const macroAnalysis = insights.analysis ? insights.analysis.macro : "No Macro Data";
    const portAnalysis = insights.analysis ? insights.analysis.portfolio : "No Portfolio Data";
    const sentimentLabel = insights.sentiment ? insights.sentiment.label : "N/A";
    const sentimentScore = insights.sentiment ? insights.sentiment.score : 0;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px;">
        <h2 style="color: #2c3e50;">🧠 Market Insights Report</h2>
        <p><strong>Date:</strong> ${today.toLocaleDateString('it-IT')}</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-left: 5px solid #007bff; margin: 20px 0;">
          <h3 style="margin-top: 0;">🚦 Market Sentiment</h3>
          <p style="font-size: 18px; font-weight: bold;">
            ${sentimentLabel} <span style="color: #666;">(${sentimentScore}/10)</span>
          </p>
        </div>

        <h3 style="border-bottom: 1px solid #ddd; padding-bottom: 5px;">🌍 Macro Analysis</h3>
        <p style="line-height: 1.6;">${macroAnalysis}</p>

        <h3 style="border-bottom: 1px solid #ddd; padding-bottom: 5px; margin-top: 20px;">💼 Portfolio Strategy</h3>
        <p style="line-height: 1.6;">${portAnalysis}</p>

        <hr style="margin-top: 30px; border: 0; border-top: 1px solid #eee;">
        <p style="font-size: 12px; color: #999; text-align: center;">
          Generated automatically by Wealth App AI
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: recipient,
      subject: `📉 Market Insights - ${today.toLocaleDateString('it-IT')}`,
      htmlBody: htmlBody
    });

    console.log("Nightly report sent successfully.");

  } catch (e) {
    console.error("Failed to send nightly report: " + e.toString());
    MailApp.sendEmail({
      to: recipient,
      subject: "⚠️ Error: Market Report Failed",
      body: "The nightly generation failed. Error: " + e.toString()
    });
  }
}

/**
 * Performs a deep-dive "Chief Risk Officer" assessment.
 * UPGRADE: Enhanced Analytical Prompt for Asset Allocation & Correlation.
 * Checks for "Fake Diversification" and specific hedging strategies.
 * Now uses Native JSON Mode for 100% parsing reliability.
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

  // 3. THE "ELITE CRO" PROMPT (XML STRUCTURED)
  const prompt = `
<role>
You are the Chief Risk Officer (CRO) of a Top-Tier Multi-Family Office.
Your methodology is based on Ray Dalio's "All Weather" principles and Taleb's Risk Management.
</role>

<portfolio_structure>
- 💵 CASH / LIQUIDITY: ${dash.summary.cash.percent} (Value: ${dash.liquidNetWorth})
- 📈 EQUITY (Stocks): ${dash.summary.stocks.percent}
- 📉 ETFS (Passive): ${dash.summary.etfs.percent}
- ⚡ CRYPTO (High Vol): ${dash.summary.crypto.percent}
</portfolio_structure>

<deep_exposure_data>
- Sector Dominance: [${topSectors}]
- Geographic Bias: [${topCountries}]
- Top Positions: ${topStocks}
- Crypto Holdings: ${topCrypto}
</deep_exposure_data>

<mission>
Conduct a forensic analysis of the Asset Allocation quality. Don't just read the numbers; interpret the CORRELATIONS.
</mission>

<key_analytical_tasks>
1. "FAKE DIVERSIFICATION" CHECK:
   - Do I own different names that act the same? (e.g. Tech Stocks + Nasdaq ETF + Crypto = 100% Correlation).
   - Identify the "Single Point of Failure" (The one factor that kills the portfolio).

2. EFFICIENCY & SIZING:
   - Is the Cash drag too high given inflation?
   - Is the Crypto allocation reckless (>5-10%) or strategic?
   - Is there Home Country Bias?

3. STRESS TEST SIMULATION:
   - Calculate expected drawdown based on the weight of High-Beta assets (Crypto/Tech) vs Low-Beta (Cash/Bonds).
</key_analytical_tasks>

<output_guidelines>
- Output strictly in JSON format.
- Keep all JSON keys in English.
- Write all JSON string values in Italian.
</output_guidelines>

<output_format>
{
  "riskScore": "1-100", 
  "riskLevel": "Basso/Medio/Alto/Estremo",
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
</output_format>
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
      
      const payload = { 
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: "application/json" } // CRITICAL: Force JSON mode
      };
      
      const res = UrlFetchApp.fetch(url, {
        method: "post", 
        contentType: "application/json", 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true
      });
      
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        
        // No regex cleanup needed, parse safely directly
        finalResult = JSON.parse(text);
        
        // Save to cache for 12 hours
        cache.put(CACHE_KEY, JSON.stringify(finalResult), 43200); 
        break; 
      }
    } catch (e) { 
      console.error("Risk AI Error: " + e.toString()); 
    }
  }

  return finalResult;
}
/**
 * Estimates annual dividend income using AI.
 * Cleans European number formats, fetches current dividend yields/payment months via Gemini,
 * and calculates the total projected yearly return for the top 15 assets.
 * Now uses Native JSON Mode for 100% parsing reliability.
 *
 * @return {Object} Total yearly estimate and a detailed list of paying assets.
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
  
  // Corrected English Prompt with XML tags for strict JSON schema enforcement
  const prompt = `
<role>You are an expert Financial Data Analyst.</role>
<task>
Estimate the current Annual Dividend Yield (as a decimal) and the Next Payment Month for the following financial assets:
[${assetString}]
</task>
<rules>
- Return strictly a JSON array of objects.
- If an asset does NOT pay a dividend (e.g., Growth stocks, Bitcoin), set "y" to 0 and "m" to "N/A".
- The "m" (month) value MUST be written in Italian (e.g., "Maggio", "Giugno").
- Use English for the JSON keys ("t", "y", "m").
</rules>
<output_format>
[
  {
    "t": "TICKER",
    "y": 0.045, 
    "m": "Mese in Italiano"
  }
]
</output_format>
  `;

  let aiData = [];

  for (let m=0; m<MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      
      const payload = { 
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: "application/json" } // CRITICAL: Force JSON mode
      };
      
      const options = { 
        method: "post", 
        contentType: "application/json", 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      };
      
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        
        // Native JSON parsing without string manipulation
        aiData = JSON.parse(text);
        break; 
      }
    } catch(e) { console.warn("Dividend AI Model Error:", e); }
  }

  let totalYearly = 0;
  let finalItems = [];

  assetList.forEach(myAsset => {
      // Find matching data from AI response (safe check on 't')
      const info = aiData.find(d => d.t && d.t.toUpperCase() === myAsset.ticker);
      
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

    OUTPUT LANGUAGE: Italian (Strictly).
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