/**
 * Fetches real-time news from MULTIPLE RSS feeds concurrently based on category.
 * Aggregates Italian and International sources, extracts images, and sorts by date.
 */
function getMarketNews(category = 'market') {
  try {
    let portfolioUrl = '';
    let allTickers = [];
    
    // Dynamically build the Portfolio URL if requested
    if (category === 'portfolio') {
      const portfolio = getLivePortfolio();
      
      if (portfolio.stocks) portfolio.stocks.forEach(item => { if (item.t) allTickers.push(item.t); });
      if (portfolio.etfs) portfolio.etfs.forEach(item => { if (item.t) allTickers.push(item.t); });
      
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

   // --- MULTI-SOURCE FEED AGGREGATOR (SAFE VERSION) ---
    const feeds = {
      'market': [
        { url: 'https://feeds.a.dj.com/rss/RSSMarketsMain.xml', source: 'Wall Street Journal' },
        { url: 'https://www.ilsole24ore.com/rss/finanza.xml', source: 'Il Sole 24 Ore' },
        { url: 'https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=10001147', source: 'CNBC' },
        { url: 'https://www.teleborsa.it/rss/news.xml', source: 'Teleborsa' },
        { url: 'https://www.milanofinanza.it/rss/news', source: 'Milano Finanza' },
        { url: 'https://finance.yahoo.com/news/rssindex', source: 'Yahoo Finance News' },
        { url: 'https://www.investing.com/rss/news_25.rss', source: 'Investing.com' }
      ],
      'portfolio': [
        { url: portfolioUrl, source: 'Yahoo Finance' }
      ],
      'world': [
        { url: 'http://feeds.bbci.co.uk/news/world/rss.xml', source: 'BBC Global' },
        { url: 'https://www.ansa.it/sito/notizie/mondo/mondo_rss.xml', source: 'ANSA Mondo' },
        { url: 'https://www.theguardian.com/world/rss', source: 'The Guardian' },
        { url: 'http://rss.cnn.com/rss/edition_world.rss', source: 'CNN World' },
        { url: 'https://www.aljazeera.com/xml/rss/all.xml', source: 'Al Jazeera' },
        { url: 'https://www.ilpost.it/feed/', source: 'Il Post' }
      ],
      'politics': [
        { url: 'http://feeds.bbci.co.uk/news/politics/rss.xml', source: 'BBC Politics' },
        { url: 'https://www.ansa.it/sito/notizie/politica/politica_rss.xml', source: 'ANSA Politica' },
        { url: 'https://www.repubblica.it/rss/politica/rss2.0.xml', source: 'Repubblica Politica' },
        { url: 'https://www.politico.eu/feed/', source: 'Politico Europe' },
        { url: 'https://www.ilfattoquotidiano.it/feed/', source: 'Il Fatto Quotidiano' }
      ],
      'science': [ 
        { url: 'https://techcrunch.com/feed/', source: 'TechCrunch' },
        { url: 'https://www.theverge.com/rss/index.xml', source: 'The Verge' },
        { url: 'https://www.ansa.it/sito/notizie/tecnologia/tecnologia_rss.xml', source: 'ANSA Tech' },
        { url: 'https://www.wired.it/feed/rss', source: 'Wired IT' },
        { url: 'https://www.sciencedaily.com/rss/all.xml', source: 'Science Daily' }
      ],
      'culture': [
        { url: 'http://feeds.bbci.co.uk/news/entertainment_and_arts/rss.xml', source: 'BBC Culture' },
        { url: 'https://www.ilpost.it/cultura/feed/', source: 'Il Post Cultura' },
        { url: 'https://variety.com/feed/', source: 'Variety' },
        { url: 'https://www.repubblica.it/rss/spettacoli_e_cultura/rss2.0.xml', source: 'Repubblica Spettacoli' }
      ]
    };

    const selectedFeeds = feeds[category] || feeds['market'];
    
    const requests = selectedFeeds.map(feed => ({
      url: feed.url,
      method: 'get',
      muteHttpExceptions: true
    }));

    let responses = [];
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      console.error("FetchAll failed totally:", e);
      // Se c'è un errore gravissimo (es. Google Server offline), ritorna lista vuota e sgancia il frontend
      return { articles: [], tickers: allTickers };
    }

    let newsList = [];
    const mediaNamespace = XmlService.getNamespace('media', 'http://search.yahoo.com/mrss/');

    responses.forEach((response, index) => {
      // Usiamo un Try-Catch interno: se un sito fallisce, salta semplicemente al prossimo senza bloccare tutto
      try {
        if (!response || response.getResponseCode() !== 200) return;
        
        const sourceName = selectedFeeds[index].source;
        const xml = response.getContentText();
        
        const document = XmlService.parse(xml);
        const root = document.getRootElement();
        const channel = root.getChild('channel');
        if (!channel) return;

        const items = channel.getChildren('item');
        const maxItems = Math.min(items.length, 30); 
        
        for (let i = 0; i < maxItems; i++) {
          const title = items[i].getChildText('title') || "No Title";
          const link = items[i].getChildText('link') || "#";
          
          let description = items[i].getChildText('description') || "";
          let rawDescription = description; 
          
          description = description.replace(/(<([^>]+)>)/gi, "").trim(); 
          description = description.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&'); 
          
          const pubDateRaw = items[i].getChildText('pubDate');
          const dateObj = new Date(pubDateRaw);
          const isoDate = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : new Date().toISOString();
          
          let imageUrl = null;
          const mediaContent = items[i].getChild('content', mediaNamespace);
          if (mediaContent && mediaContent.getAttribute('url')) imageUrl = mediaContent.getAttribute('url').getValue();
          
          if (!imageUrl) {
              const mediaThumbnail = items[i].getChild('thumbnail', mediaNamespace);
              if (mediaThumbnail && mediaThumbnail.getAttribute('url')) imageUrl = mediaThumbnail.getAttribute('url').getValue();
          }
          if (!imageUrl) {
              const enclosure = items[i].getChild('enclosure');
              if (enclosure && enclosure.getAttribute('type') && enclosure.getAttribute('type').getValue().startsWith('image')) {
                  imageUrl = enclosure.getAttribute('url').getValue();
              }
          }
          if (!imageUrl && rawDescription) {
              const imgMatch = rawDescription.match(/<img[^>]+src="([^">]+)"/);
              if (imgMatch) imageUrl = imgMatch[1];
          }
          if (imageUrl && imageUrl.startsWith('http://')) {
              imageUrl = imageUrl.replace('http://', 'https://');
          }
          
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
          
          newsList.push({
            title: title,
            summary: description,
            link: link,
            date: isoDate,
            source: sourceName,
            category: category,
            relatedTickers: articleTickers,
            imageUrl: imageUrl
          });
        }
      } catch (e) {
        // Se c'è un errore nell'XML di un sito specifico, loggalo e vai oltre
        console.error("Errore lettura feed per: " + selectedFeeds[index].source, e);
      }
    });

    newsList.sort((a, b) => new Date(b.date) - new Date(a.date));
    newsList = newsList.slice(0, 100);
    
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

/**
 * Fetches saved news from a dedicated Google Sheet ('Saved_News').
 * Creates the sheet automatically if it doesn't exist yet.
 * @returns {Array<Object>} List of saved articles.
 */
function getCloudSavedNews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Saved_News');
  
  if (!sheet) {
    sheet = ss.insertSheet('Saved_News');
    sheet.appendRow(['Link', 'Title', 'Summary', 'Date', 'Source', 'Category', 'ImageUrl', 'RelatedTickers']);
    sheet.setFrozenRows(1);
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3f4f6");
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const savedNews = [];
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    savedNews.push({
      link: row[0],
      title: row[1],
      summary: row[2],
      date: row[3],
      source: row[4],
      category: row[5],
      imageUrl: row[6],
      relatedTickers: row[7] ? JSON.parse(row[7]) : []
    });
  }
  // Return reversed to show the newest saved at the top
  return savedNews.reverse();
}

/**
 * Toggles an article's saved status in the 'Saved_News' sheet.
 * @param {Object} article - The news article object.
 * @returns {boolean} True if added, false if removed.
 */
function toggleCloudSavedNews(article) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Saved_News');
  if (!sheet) {
      sheet = ss.insertSheet('Saved_News');
      sheet.appendRow(['Link', 'Title', 'Summary', 'Date', 'Source', 'Category', 'ImageUrl', 'RelatedTickers']);
  }

  const data = sheet.getDataRange().getValues();
  let foundIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
      if (data[i][0] === article.link) {
          foundIndex = i + 1; // arrays are 0-indexed, sheet rows are 1-indexed
          break;
      }
  }

  if (foundIndex !== -1) {
      sheet.deleteRow(foundIndex);
      return false; // Removed
  } else {
      sheet.appendRow([
          article.link,
          article.title,
          article.summary,
          article.date,
          article.source,
          article.category,
          article.imageUrl || '',
          JSON.stringify(article.relatedTickers || [])
      ]);
      return true; // Added
  }
}