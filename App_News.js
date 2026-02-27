/**
 * Fetches real-time news from various RSS feeds based on category.
 * Dynamically loads portfolio tickers if category is 'portfolio'.
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
      
      // Extract Stocks and ETFs
      if (portfolio.stocks) {
        portfolio.stocks.forEach(item => { if (item.t) allTickers.push(item.t); });
      }
      if (portfolio.etfs) {
        portfolio.etfs.forEach(item => { if (item.t) allTickers.push(item.t); });
      }
      
      // Extract Crypto and format for Yahoo Finance (e.g., BTC -> BTC-USD)
      if (portfolio.crypto) {
        portfolio.crypto.forEach(item => {
          if (item.t) {
            let ticker = item.t;
            if (!ticker.includes('-')) {
              ticker = ticker + '-USD';
            }
            allTickers.push(ticker);
          }
        });
      }
      
      // Remove duplicates and limit to top 15 to prevent URL length errors
      allTickers = [...new Set(allTickers)].slice(0, 15);
      
      if (allTickers.length > 0) {
        const tickerString = encodeURIComponent(allTickers.join(','));
        portfolioUrl = `https://feeds.finance.yahoo.com/rss/2.0/headline?s=${tickerString}&region=US&lang=en-US`;
      } else {
        // Fallback to S&P 500 if portfolio is empty or fails to load
        portfolioUrl = 'https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC&region=US&lang=en-US';
      }
    }

    // Define RSS feeds for different categories
    const feeds = {
      'market': 'https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC,%5EDJI,BTC-USD&region=US&lang=en-US',
      'portfolio': portfolioUrl,
      'world': 'http://feeds.bbci.co.uk/news/world/rss.xml',
      'politics': 'http://feeds.bbci.co.uk/news/politics/rss.xml',
      'science': 'http://feeds.bbci.co.uk/news/science_and_environment/rss.xml',
      'culture': 'http://feeds.bbci.co.uk/news/entertainment_and_arts/rss.xml'
    };

    const url = feeds[category] || feeds['market'];
    
    const options = {
      'method': 'get',
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      console.error("Failed to fetch RSS feed: HTTP " + response.getResponseCode());
      return { articles: [], tickers: allTickers };
    }

    const xml = response.getContentText();
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const channel = root.getChild('channel');
    
    if (!channel) {
      console.error("Invalid XML structure: missing channel node.");
      return { articles: [], tickers: allTickers };
    }

    const items = channel.getChildren('item');
    const newsList = [];
    
    const maxItems = Math.min(items.length, 15); 
    
    let sourceName = "Yahoo Finance";
    if (category !== 'market' && category !== 'portfolio') {
      sourceName = "BBC Global";
    }
    
    for (let i = 0; i < maxItems; i++) {
      const title = items[i].getChildText('title') || "No Title";
      const link = items[i].getChildText('link') || "#";
      
      const pubDateRaw = items[i].getChildText('pubDate');
      const dateObj = new Date(pubDateRaw);
      const isoDate = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : new Date().toISOString();
      
      newsList.push({
        title: title,
        link: link,
        date: isoDate,
        source: sourceName,
        category: category
      });
    }
    
    return { articles: newsList, tickers: allTickers };
    
  } catch (error) {
    console.error("Error fetching news for category " + category + ": " + error.toString());
    return { articles: [], tickers: [] };
  }
}