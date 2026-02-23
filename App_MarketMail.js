/**
 * Identifies upcoming earnings and macro events using AI for the next 14 days,
 * and sends calendar invites (.ics) to the specified email.
 *
 * @param {string} targetEmail - The email address to send the calendars to.
 * @return {Object} Status of the operation and generated events.
 */
function sendMarketEventsCalendarToOutlook(targetEmail = "alessandro.saladino01@gmail.com") {
  const PROPERTY_KEY = "MARKET_EVENTS_SYNC_V1";
  const props = PropertiesService.getScriptProperties();
  
  // 1. Fetch portfolio and parse assets to get tickers for earnings
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
    .slice(0, 30); // Expanded to 30 to cover more portfolio earnings

  if (assetList.length === 0) return { status: "No assets found", items: [] };

  const assetString = assetList.map(a => a.ticker).join(", ");
  const API_KEY = GEMINI_API_KEY; 
  const MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"];
  
  // 2. Define Time Window for the upcoming 14 days
  const today = new Date();
  const twoWeeks = new Date();
  twoWeeks.setDate(today.getDate() + 14);
  
  const formatDate = (d) => {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  
  const todayStr = formatDate(today);
  const twoWeeksStr = formatDate(twoWeeks);

  // 3. Prompt targeted strictly at Earnings and Macro events
  const prompt = `
<role>You are an expert Institutional Hedge Fund Macro Strategist.</role>
<task>
Today's date is ${todayStr}. Identify the most critical upcoming MACRO EVENTS (e.g., CPI, Fed/ECB Rate Decisions, NFP) and EARNINGS RELEASES for the following tickers:
[${assetString}]
</task>
<rules>
- Strictly look for events scheduled between ${todayStr} and ${twoWeeksStr}.
- Return strictly a JSON array of objects.
- "type" MUST be either "MACRO" or "EARNINGS".
- The "d" (date) value MUST be formatted EXACTLY as YYYY-MM-DD.
- Provide a brief analytic context in the "desc" field (in Italian).
- Use English for the JSON keys ("type", "title", "desc", "d", "ticker").
- If no significant events are found in this 14-day window, return an empty array [].
</rules>
<output_format>
[
  {
    "type": "MACRO",
    "ticker": "N/A",
    "title": "US CPI Data Release",
    "desc": "Dato cruciale per capire le prossime mosse della FED sui tassi.",
    "d": "2024-05-15"
  },
  {
    "type": "EARNINGS",
    "ticker": "AAPL",
    "title": "Apple Q3 Earnings",
    "desc": "Focus sulle vendite iPhone e guidance per il prossimo trimestre.",
    "d": "2024-05-18"
  }
]
</output_format>
  `;

  let aiData = [];

  // 4. API Call to Gemini
  for (let m=0; m < MODELS.length; m++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELS[m]}:generateContent?key=${API_KEY}`;
      const payload = { 
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: "application/json" }
      };
      
      const options = { 
        method: "post", contentType: "application/json", 
        payload: JSON.stringify(payload), muteHttpExceptions: true 
      };
      
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
        aiData = JSON.parse(text);
        break; 
      }
    } catch(e) { console.warn("Market Events AI Model Error:", e); }
  }

  let attachments = [];
  let generatedEvents = [];

  // 5. Generate .ics calendar files ONLY for events strictly within the next 14 days
  if (Array.isArray(aiData)) {
      aiData.forEach(event => {
          if (event.d && event.d >= todayStr && event.d <= twoWeeksStr) {
              
              const dateParts = event.d.split("-");
              if(dateParts.length === 3) {
                  const year = dateParts[0];
                  const month = dateParts[1];
                  const day = dateParts[2];
                  
                  const icsDate = `${year}${month}${day}`;
                  
                  // Dynamic subject icon based on event type
                  const icon = event.type === "EARNINGS" ? "📊" : "🏛️";
                  const subject = `${icon} ${event.type}: ${event.title}`;
                  const description = `${event.desc}\\n\\nType: ${event.type}\\nRelated Asset: ${event.ticker}`;
                  
                  // Generate UID and DTSTAMP required by Outlook's strict policies
                  const uid = Utilities.getUuid() + "@wealthapp.local";
                  const nowStr = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + "Z";

                  const icsContent = [
                    "BEGIN:VCALENDAR",
                    "VERSION:2.0",
                    "PRODID:-//WealthApp//MarketEvents//EN",
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

                  // Create a safe filename
                  const safeTitle = event.title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
                  
                  attachments.push({
                    fileName: `${event.d}_${event.type}_${safeTitle}.ics`,
                    mimeType: "text/calendar",
                    content: icsContent
                  });
                  
                  generatedEvents.push(event);
              }
          }
      });
  }

  // 6. Send Email if there are upcoming events
  if (attachments.length > 0) {
      const blobs = attachments.map(att => 
         Utilities.newBlob(att.content, att.mimeType, att.fileName)
      );

      MailApp.sendEmail({
        to: targetEmail,
        subject: "🔔 Upcoming Market Events & Earnings (Next 14 Days)",
        body: "Attached are the estimated Macro events and Earnings for your portfolio over the next 2 weeks. Open the .ics files to add them to your calendar.",
        attachments: blobs
      });
      
      props.setProperty(PROPERTY_KEY, Date.now().toString());
  } else {
      console.log("No market events or earnings scheduled for the next 14 days.");
  }

  return { status: "Process completed", events: generatedEvents };
}

/**
 * Sets up a weekly time-driven trigger to run the Market Events sync every Sunday.
 * RUN THIS FUNCTION ONCE MANUALLY from the Apps Script Editor to activate it.
 */
function setupWeeklyMarketEventsTrigger() {
  const functionName = "sendMarketEventsCalendarToOutlook";
  
  // Remove existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger: Every Sunday at 10 AM (Apps Script time) - staggered from the dividends script
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(10)
    .create();
    
  Logger.log("Weekly trigger successfully set for Sunday at 10:00 AM.");
}