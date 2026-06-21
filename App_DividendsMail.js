/**
 * Estimates next dividend payment dates using AI, filters for the upcoming week,
 * and sends calendar invites (.ics) to the specified email.
 * Uses Universal AI Router for robust fallback (OpenRouter -> Gemini).
 *
 * @param {string} targetEmail - The email address to send the calendars to.
 * @return {Object} Status of the operation and generated events.
 */
function sendDividendCalendarToOutlook(targetEmail = getNotifyEmail()) { // FIX (2.17): email centralizzata
  const PROPERTY_KEY = "DIVIDEND_CALENDAR_SYNC_V1";
  const props = PropertiesService.getScriptProperties();
  
  // 1. Fetch portfolio and parse assets
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
    .map(a => ({ ticker: a.t.toUpperCase(), value: parseValue(a.val) }))
    .filter(a => a.value > 0) 
    .sort((a, b) => b.value - a.value)
    .slice(0, 15);

  if (assetList.length === 0) return { status: "No assets found", items: [] };

  const assetString = assetList.map(a => a.ticker).join(", ");
  
  // 2. Define Time Window for the upcoming week (Today to Today + 7 days)
  const today = new Date();
  const nextWeek = new Date();
  nextWeek.setDate(today.getDate() + 7);
  
  const formatDate = (d) => {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  
  const todayStr = formatDate(today);
  const nextWeekStr = formatDate(nextWeek);

  // 3. Prompt modified to provide current date context to the AI
  const prompt = `
<role>You are an expert Financial Data Analyst.</role>
<task>
Today's date is ${todayStr}. Estimate the current Annual Dividend Yield (as a decimal) and the EXACT Next Dividend Payment Date (YYYY-MM-DD) for the following financial assets:
[${assetString}]
</task>
<rules>
- Return strictly a JSON array of objects.
- If an asset does NOT pay a dividend (e.g., Growth stocks, Bitcoin), set "y" to 0 and "d" to "N/A".
- The "d" (date) value MUST be formatted as YYYY-MM-DD. 
- Focus especially on identifying if any payments are scheduled between ${todayStr} and ${nextWeekStr}.
- Use English for the JSON keys ("t", "y", "d").
</rules>
<output_format>
[
  {
    "t": "TICKER",
    "y": 0.045, 
    "d": "2024-05-15"
  }
]
</output_format>
  `;

  let aiData = [];

  // 4. API Call using Universal AI Router
  console.log("Dividend Calendar AI: Attempting OPENROUTER...");
  let responseText = fetchUniversalAI(prompt, 'OPENROUTER');

  if (!responseText) {
    console.log("Dividend Calendar AI: OPENROUTER failed. Falling back to GEMINI...");
    responseText = fetchUniversalAI(prompt, 'GEMINI');
  }

  // Parse the AI response robustly
  if (responseText) {
    try {
      let cleanText = responseText.replace(/```json/gi, "").replace(/```/g, "").trim();
      
      const firstBracket = cleanText.indexOf('[');
      const lastBracket = cleanText.lastIndexOf(']');
      if (firstBracket !== -1 && lastBracket !== -1) {
          cleanText = cleanText.substring(firstBracket, lastBracket + 1);
          aiData = JSON.parse(cleanText);
      }
    } catch(e) { 
      console.warn("Dividend Calendar JSON Parsing Error:", e); 
    }
  }

  let attachments = [];
  let generatedEvents = [];

  // 5. Generate .ics calendar files ONLY for events within the next 7 days
  if (Array.isArray(aiData)) {
    assetList.forEach(myAsset => {
        const info = aiData.find(d => d.t && d.t.toUpperCase() === myAsset.ticker);
        
        if (info && info.y > 0 && info.d && info.d !== "N/A") {
            // Strictly filter dates to ensure they fall within the upcoming week
            if (info.d >= todayStr && info.d <= nextWeekStr) {
                let estimatedYearly = myAsset.value * info.y;
                let estPayment = (estimatedYearly / 4).toFixed(2); 
                
                const dateParts = info.d.split("-");
                if(dateParts.length === 3) {
                    const year = dateParts[0];
                    const month = dateParts[1];
                    const day = dateParts[2];
                    
                    const icsDate = `${year}${month}${day}`;
                    
                    const subject = `💰 Dividend Payment: ${myAsset.ticker}`;
                    const description = `Estimated dividend payment for ${myAsset.ticker}.\\nEstimated Amount: ~€${estPayment}\\nYearly Yield: ${(info.y * 100).toFixed(2)}%\\nTotal Asset Value: €${myAsset.value}`;
                    
                    // Generate UID and DTSTAMP required by Outlook's strict policies
                    const uid = Utilities.getUuid() + "@wealthapp.local";
                    const nowStr = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + "Z";

                    const icsContent = [
                      "BEGIN:VCALENDAR",
                      "VERSION:2.0",
                      "PRODID:-//WealthApp//DividendTracker//EN",
                      "METHOD:PUBLISH",
                      "BEGIN:VEVENT",
                      `UID:${uid}`,
                      `DTSTAMP:${nowStr}`,
                      `DTSTART;VALUE=DATE:${icsDate}`,
                      `DTEND;VALUE=DATE:${icsDate}`,
                      `SUMMARY:${subject}`,
                      `DESCRIPTION:${description}`,
                      "STATUS:CONFIRMED",
                      "SEQUENCE:0",
                      "END:VEVENT",
                      "END:VCALENDAR"
                    ].join("\r\n");

                    attachments.push({
                      fileName: `${myAsset.ticker}_dividend_${info.d}.ics`,
                      mimeType: "text/calendar",
                      content: icsContent
                    });
                    
                    generatedEvents.push({ ticker: myAsset.ticker, date: info.d, amount: estPayment });
                }
            }
        }
    });
  }

  // 6. Send Email if there are upcoming dividends
  if (attachments.length > 0) {
      const blobs = attachments.map(att => 
         Utilities.newBlob(att.content, att.mimeType, att.fileName)
      );

      MailApp.sendEmail({
        to: targetEmail,
        subject: "📅 Upcoming Week Dividend Payments",
        body: "Attached are your estimated dividend payments for the upcoming 7 days. Open the .ics files to add them to your calendar.",
        attachments: blobs
      });
      
      props.setProperty(PROPERTY_KEY, Date.now().toString());
  } else {
      console.log("No dividends scheduled for the next 7 days.");
  }

  return { status: "Process completed", events: generatedEvents };
}

/**
 * Sets up a weekly time-driven trigger to run the dividend sync every Sunday.
 * RUN THIS FUNCTION ONCE MANUALLY from the Apps Script Editor to activate it.
 */
function setupWeeklyDividendSyncTrigger() {
  const functionName = "sendDividendCalendarToOutlook";
  
  // Remove existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger: Every Sunday at 9 AM (Apps Script time)
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();
    
  Logger.log("Weekly trigger successfully set for Sunday at 9:00 AM.");
}

/**
 * TEST FUNCTION: Estimates next dividend payment dates using AI for the next 365 days
 * just to test the email delivery and .ics formatting.
 */
function test_sendDividendCalendarToOutlook() {
  const targetEmail = getNotifyEmail(); // FIX (2.17): email centralizzata
  
  // 1. Fetch portfolio and parse assets
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
    .map(a => ({ ticker: a.t.toUpperCase(), value: parseValue(a.val) }))
    .filter(a => a.value > 0) 
    .sort((a, b) => b.value - a.value)
    .slice(0, 15);

  if (assetList.length === 0) {
    console.log("Nessun asset trovato nel portafoglio.");
    return;
  }

  const assetString = assetList.map(a => a.ticker).join(", ");
  
  // WIDENED TIME WINDOW FOR TESTING: Next 365 days
  const today = new Date();
  const nextYear = new Date();
  nextYear.setDate(today.getDate() + 365);
  
  const formatDate = (d) => {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  
  const todayStr = formatDate(today);
  const nextYearStr = formatDate(nextYear);

  const prompt = `
<role>You are an expert Financial Data Analyst.</role>
<task>
Today's date is ${todayStr}. Estimate the current Annual Dividend Yield (as a decimal) and the EXACT Next Dividend Payment Date (YYYY-MM-DD) for the following financial assets:
[${assetString}]
</task>
<rules>
- Return strictly a JSON array of objects.
- If an asset does NOT pay a dividend, set "y" to 0 and "d" to "N/A".
- The "d" (date) value MUST be formatted as YYYY-MM-DD. 
- Use English for the JSON keys ("t", "y", "d").
</rules>
<output_format>
[
  {
    "t": "TICKER",
    "y": 0.045, 
    "d": "2024-05-15"
  }
]
</output_format>
  `;

  let aiData = [];

  console.log("TEST Dividend AI: Attempting OPENROUTER...");
  let responseText = fetchUniversalAI(prompt, 'OPENROUTER');

  if (!responseText) {
    console.log("TEST Dividend AI: OPENROUTER failed. Falling back to GEMINI...");
    responseText = fetchUniversalAI(prompt, 'GEMINI');
  }

  if (responseText) {
    try {
      let cleanText = responseText.replace(/```json/gi, "").replace(/```/g, "").trim();
      
      const firstBracket = cleanText.indexOf('[');
      const lastBracket = cleanText.lastIndexOf(']');
      if (firstBracket !== -1 && lastBracket !== -1) {
          cleanText = cleanText.substring(firstBracket, lastBracket + 1);
          aiData = JSON.parse(cleanText);
      }
    } catch(e) { 
      console.warn("TEST Dividend JSON Parsing Error:", e); 
    }
  }

  let attachments = [];

  if (Array.isArray(aiData)) {
    assetList.forEach(myAsset => {
        const info = aiData.find(d => d.t && d.t.toUpperCase() === myAsset.ticker);
        
        if (info && info.y > 0 && info.d && info.d !== "N/A") {
            // Broad filter just to guarantee some emails for the test
            if (info.d >= todayStr && info.d <= nextYearStr) {
                let estimatedYearly = myAsset.value * info.y;
                let estPayment = (estimatedYearly / 4).toFixed(2); 
                
                const dateParts = info.d.split("-");
                if(dateParts.length === 3) {
                    const year = dateParts[0];
                    const month = dateParts[1];
                    const day = dateParts[2];
                    
                    const icsDate = `${year}${month}${day}`;
                    
                    const subject = `💰 TEST Dividend: ${myAsset.ticker}`;
                    const description = `TEST EMAIL.\\nEstimated Amount: ~€${estPayment}\\nYearly Yield: ${(info.y * 100).toFixed(2)}%`;
                    
                    // Generate UID and DTSTAMP required by Outlook's strict policies
                    const uid = Utilities.getUuid() + "@wealthapp.local";
                    const nowStr = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + "Z";

                    const icsContent = [
                      "BEGIN:VCALENDAR",
                      "VERSION:2.0",
                      "PRODID:-//WealthApp//DividendTracker//EN",
                      "METHOD:PUBLISH",
                      "BEGIN:VEVENT",
                      `UID:${uid}`,
                      `DTSTAMP:${nowStr}`,
                      `DTSTART;VALUE=DATE:${icsDate}`,
                      `DTEND;VALUE=DATE:${icsDate}`,
                      `SUMMARY:${subject}`,
                      `DESCRIPTION:${description}`,
                      "STATUS:CONFIRMED",
                      "SEQUENCE:0",
                      "END:VEVENT",
                      "END:VCALENDAR"
                    ].join("\r\n");

                    attachments.push({
                      fileName: `TEST_${myAsset.ticker}_dividend.ics`,
                      mimeType: "text/calendar",
                      content: icsContent
                    });
                }
            }
        }
    });
  }

  if (attachments.length > 0) {
      const blobs = attachments.map(att => 
         Utilities.newBlob(att.content, att.mimeType, att.fileName)
      );

      MailApp.sendEmail({
        to: targetEmail,
        subject: "🔧 TEST: Calendario Dividendi",
        body: "Questo è un test per verificare la ricezione dei file .ics di Outlook.",
        attachments: blobs
      });
      console.log("Test inviato con successo! Controlla la casella email.");
  } else {
      console.log("Nessun dividendo trovato neanche allargando la ricerca a 365 giorni.");
  }
}