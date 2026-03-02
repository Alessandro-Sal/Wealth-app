/**
 * Fetches real-time news from various RSS feeds based on category.
 * NOW INCLUDES: Image Extraction from XML feeds.
 * @param {string} category - The category of news to fetch.
 * @returns {Object} An object containing the articles array and the used tickers.
 */
function getMarketNews(category = 'market') {
  try {
    let portfolioUrl = '';
    let allTickers = [];
    
    // Dynamically build the Portfolio URL if requested
    if (category === 'portfolio') {
      const portfolio = getLivePortfolio();
      
      if (portfolio.stocks) {
        portfolio.stocks.forEach(item => { if (item.t) allTickers.push(item.t); });
      }
      if (portfolio.etfs) {
        portfolio.etfs.forEach(item => { if (item.t) allTickers.push(item.t); });
      }
      
      if (portfolio.crypto) {
        portfolio.crypto.forEach(item => {
          if (item.t) {
            let ticker = item.t;
            if (!ticker.includes('-')) { ticker = ticker + '-USD'; }
            allTickers.push(ticker);
          }
        });
      }
      
      allTickers = [...new Set(allTickers)].slice(0, 15);
      
      if (allTickers.length > 0) {
        const tickerString = encodeURIComponent(allTickers.join(','));
        portfolioUrl = `https://feeds.finance.yahoo.com/rss/2.0/headline?s=${tickerString}&region=US&lang=en-US`;
      } else {
        portfolioUrl = 'https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC&region=US&lang=en-US';
      }
    }

    const feeds = {
      'market': 'https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC,%5EDJI,BTC-USD&region=US&lang=en-US',
      'portfolio': portfolioUrl,
      'world': 'http://feeds.bbci.co.uk/news/world/rss.xml',
      'politics': 'http://feeds.bbci.co.uk/news/politics/rss.xml',
      'science': 'http://feeds.bbci.co.uk/news/science_and_environment/rss.xml',
      'culture': 'http://feeds.bbci.co.uk/news/entertainment_and_arts/rss.xml'
    };

    const url = feeds[category] || feeds['market'];
    const options = { 'method': 'get', 'muteHttpExceptions': true };
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      console.error("Failed to fetch RSS feed");
      return { articles: [], tickers: allTickers };
    }

    const xml = response.getContentText();
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const channel = root.getChild('channel');
    
    if (!channel) return { articles: [], tickers: allTickers };

    const items = channel.getChildren('item');
    const newsList = [];
    
    // Download up to 50 items for pagination
    const maxItems = Math.min(items.length, 50); 
    
    let sourceName = "Yahoo Finance";
    if (category !== 'market' && category !== 'portfolio') {
      sourceName = "BBC Global";
    }

    // Set up XML namespaces for media parsing
    const mediaNamespace = XmlService.getNamespace('media', 'http://search.yahoo.com/mrss/');
    
    for (let i = 0; i < maxItems; i++) {
      const title = items[i].getChildText('title') || "No Title";
      const link = items[i].getChildText('link') || "#";
      
      let description = items[i].getChildText('description') || "";
      let rawDescription = description; // Save raw to extract potential img tags later
      
      description = description.replace(/(<([^>]+)>)/gi, "").trim(); 
      description = description.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&'); 
      
      const pubDateRaw = items[i].getChildText('pubDate');
      const dateObj = new Date(pubDateRaw);
      const isoDate = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : new Date().toISOString();
      
      // --- IMAGE EXTRACTION LOGIC ---
      let imageUrl = null;
      
      // 1. Try <media:content> (Standard Yahoo Format)
      const mediaContent = items[i].getChild('content', mediaNamespace);
      if (mediaContent && mediaContent.getAttribute('url')) {
          imageUrl = mediaContent.getAttribute('url').getValue();
      }
      
      // 2. Try <media:thumbnail> (Standard BBC Format)
      if (!imageUrl) {
          const mediaThumbnail = items[i].getChild('thumbnail', mediaNamespace);
          if (mediaThumbnail && mediaThumbnail.getAttribute('url')) {
              imageUrl = mediaThumbnail.getAttribute('url').getValue();
          }
      }

      // 3. Try <enclosure> (Standard Podcast/Generic RSS Format)
      if (!imageUrl) {
          const enclosure = items[i].getChild('enclosure');
          if (enclosure && enclosure.getAttribute('type') && enclosure.getAttribute('type').getValue().startsWith('image')) {
              imageUrl = enclosure.getAttribute('url').getValue();
          }
      }
      
      // 4. Fallback: Parse the raw description for an HTML <img> tag
      if (!imageUrl && rawDescription) {
          const imgMatch = rawDescription.match(/<img[^>]+src="([^">]+)"/);
          if (imgMatch) imageUrl = imgMatch[1];
      }
      // -----------------------------
      
      let articleTickers = [];
      if (category === 'portfolio' && allTickers.length > 0) {
        allTickers.forEach(ticker => {
           let cleanTicker = ticker.replace('-USD', '');
           let regex = new RegExp("\\b" + cleanTicker + "\\b", "i");
           if (regex.test(title) || regex.test(description) || link.toUpperCase().includes(cleanTicker.toUpperCase())) {
              articleTickers.push(cleanTicker.toUpperCase());
           }
        });
      }
      
      // 4. Fallback: Parse the raw description for an HTML <img> tag
      if (!imageUrl && rawDescription) {
          const imgMatch = rawDescription.match(/<img[^>]+src="([^">]+)"/);
          if (imgMatch) imageUrl = imgMatch[1];
      }
      
      // --- NEW FIX: Force HTTPS to prevent Mixed Content blocking in Apps Script ---
      if (imageUrl && imageUrl.startsWith('http://')) {
          imageUrl = imageUrl.replace('http://', 'https://');
      }
      // -----------------------------
      
      let articleTickers = [];
      newsList.push({
        title: title,
        summary: description,
        link: link,
        date: isoDate,
        source: sourceName,
        category: category,
        relatedTickers: articleTickers,
        imageUrl: imageUrl // Pass the extracted image URL to frontend
      });
    }
    
    return { articles: newsList, tickers: allTickers };
    
  } catch (error) {
    console.error("News fetch error:", error);
    return { articles: [], tickers: [] };
  }
}

/**
 * Generates an AI summary of the current top news using the universal AI router.
 * @param {Array} newsList - Array of fetched news objects.
 * @returns {string} AI generated summary in Italian.
 */
function getNewsAIBriefing(newsList) {
  try {
    if (!newsList || newsList.length === 0) {
      return "No news available to summarize at the moment.";
    }
    
    // Extract top 10 titles and sources for the prompt to avoid token overflow
    const topNews = newsList.slice(0, 10).map(n => `- ${n.title} (${n.source})`).join("\n");
    
    const prompt = `Sei un analista finanziario di Wall Street. Crea un "Morning Briefing" di 3 brevi punti (massimo 400 caratteri totali) riassumendo in italiano il sentiment generale e i temi principali di queste notizie di mercato:\n\n${topNews}\n\nUsa le emoji. Non fare premesse, sii diretto e professionale.`;
    
    // Call the provided universal AI function (Defaulting to GEMINI)
    const responseText = fetchUniversalAI(prompt, 'GEMINI', false);
    
    if (responseText) {
      return responseText;
    } else {
      return "⚠️ AI service is temporarily unavailable. Please try again later.";
    }
  } catch (error) {
    console.error("Error generating AI Briefing: " + error.toString());
    return "⚠️ An error occurred while generating the summary.";
  }
}


/**
 * Generates an AI analysis/context for a single article based on its title and summary.
 * @param {string} title - The article title
 * @param {string} summary - The article snippet/description
 * @returns {string} AI generated insight in Italian
 */
function getSingleArticleAISummary(title, summary) {
  try {
    const prompt = `Sei un analista finanziario senior. Spiega brevemente questa notizia e le sue possibili implicazioni sul mercato in massimo 3 o 4 righe (in italiano). Usa un tono professionale e diretto.\n\nTitolo: ${title}\nEstratto: ${summary}`;
    
    // Using GEMINI as the default fast provider for individual news
    const responseText = fetchUniversalAI(prompt, 'GEMINI', false);
    
    if (responseText) {
      return responseText;
    } else {
      return "⚠️ AI service temporarily unavailable.";
    }
  } catch (error) {
    console.error("Error generating single article insight: " + error.toString());
    return "⚠️ An error occurred during analysis.";
  }
}