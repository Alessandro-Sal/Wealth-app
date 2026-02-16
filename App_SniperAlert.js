/**
 * App_SniperAlert.js
 * Scans LIVE portfolio for rapid price movements (> 4%).
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

  const allAssets = [...portfolio.stocks, ...portfolio.crypto];
  let alerts = [];

  // 2. Scan for Volatility
  allAssets.forEach(asset => {
    // Check daily change (dCh)
    let changeStr = asset.dCh || "0%"; 
    let change = parseFloat(changeStr.replace('%', '').replace(',', '.'));

    // THRESHOLDS: Drop < -4% or Pump > +5%
    if (change <= -4.0) {
      alerts.push(`🔻 <b>${asset.t}</b> is crashing: <span style="color:red">${changeStr}</span>`);
    } else if (change >= 5.0) {
      alerts.push(`🚀 <b>${asset.t}</b> is pumping: <span style="color:green">+${changeStr}</span>`);
    }
  });

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
      </div>`
    });
    console.log(`Sniper Alert Sent: ${alerts.length} assets.`);
  } else {
    console.log("Sniper Scan: Markets are calm.");
  }
}

/**
 * Generates quick trading advice with AI Fallback.
 */
function generateSniperAI(alerts) {
  const API_KEY = GEMINI_API_KEY;
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