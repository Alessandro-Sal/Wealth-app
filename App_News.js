/**
 * Fetches real-time market news from Yahoo Finance RSS feed.
 * Using RSS avoids the need for a dedicated API key and rate limits.
 * @returns {Array} An array of news objects containing title, link, date, and source.
 */
function getMarketNews() {
  try {
    // Top market news from Yahoo Finance (S&P 500, Dow Jones, Crypto)
    // Note: The '^' character is URL-encoded as '%5E' to prevent Apps Script "Invalid argument" errors.
    const url = 'https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC,%5EDJI,BTC-USD&region=US&lang=en-US';
    
    const options = {
      'method': 'get',
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      console.error("Failed to fetch RSS feed: HTTP " + response.getResponseCode());
      return [];
    }

    const xml = response.getContentText();
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const channel = root.getChild('channel');
    
    if (!channel) {
      console.error("Invalid XML structure: missing channel node.");
      return [];
    }

    const items = channel.getChildren('item');
    const newsList = [];
    
    // Limit the output to the latest 15 breaking news items
    const maxItems = Math.min(items.length, 15); 
    
    for (let i = 0; i < maxItems; i++) {
      const title = items[i].getChildText('title') || "No Title";
      const link = items[i].getChildText('link') || "#";
      
      // Parse the date to a standard ISO string for frontend consistency
      const pubDateRaw = items[i].getChildText('pubDate');
      const dateObj = new Date(pubDateRaw);
      const isoDate = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : new Date().toISOString();
      
      newsList.push({
        title: title,
        link: link,
        date: isoDate,
        source: "Yahoo Finance"
      });
    }
    
    return newsList;
    
  } catch (error) {
    console.error("Error fetching breaking news: " + error.toString());
    return [];
  }
}