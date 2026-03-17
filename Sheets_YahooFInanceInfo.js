/**
 * Fetches sector, industry, country, or exchange from Yahoo Finance API.
 * Maps specific countries to macro-regions (EUROPE, USA) and customizes Exchange tickers.
 * @param {string} ticker The stock symbol.
 * @param {string} type "sector", "industry", "country", or "exchange".
 * @customfunction
 */
function YAHOO_INFO(ticker, type) {
  if (!ticker) return "";
  
  var reqType = type.toLowerCase().trim();
  var cleanTicker = ticker.toString().toUpperCase().trim();
  
  // Map of Google Finance prefixes to Yahoo Finance suffixes
  var exchangeMap = {
    "BIT:": ".MI", "EPA:": ".PA", "FRA:": ".DE", "ETR:": ".DE",
    "AMS:": ".AS", "LON:": ".L",  "BME:": ".MC", "SWX:": ".SW",
    "VTX:": ".SW", "VIE:": ".VI", "STO:": ".ST"
  };
  
  for (var prefix in exchangeMap) {
    if (cleanTicker.indexOf(prefix) === 0) {
      cleanTicker = cleanTicker.replace(prefix, "") + exchangeMap[prefix];
      break;
    }
  }
  
  var cache = CacheService.getDocumentCache();
  var dataCacheKey = "INFO_" + cleanTicker + "_" + reqType;
  
  var cachedData = cache.get(dataCacheKey);
  if (cachedData) return cachedData;
  
  try {
    var cookieString = cache.get("YAHOO_COOKIE");
    var crumb = cache.get("YAHOO_CRUMB");
    
    if (!cookieString || !crumb) {
      var cookieResponse = UrlFetchApp.fetch("https://fc.yahoo.com/", {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0" }
      });
      
      var headers = cookieResponse.getAllHeaders();
      var setCookie = headers['Set-Cookie'] || headers['set-cookie'];
      cookieString = "";
      
      if (setCookie) {
        cookieString = Array.isArray(setCookie) ? setCookie[0] : setCookie;
        cookieString = cookieString.split(';')[0]; 
      }
      
      var crumbResponse = UrlFetchApp.fetch("https://query1.finance.yahoo.com/v1/test/getcrumb", {
        muteHttpExceptions: true,
        headers: { "Cookie": cookieString, "User-Agent": "Mozilla/5.0" }
      });
      crumb = crumbResponse.getContentText();
      
      if (cookieString && crumb) {
        cache.put("YAHOO_COOKIE", cookieString, 21600);
        cache.put("YAHOO_CRUMB", crumb, 21600);
      }
    }
    
    var url = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" + cleanTicker + "?modules=assetProfile,fundProfile,price&crumb=" + crumb;
    var dataResponse = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { "Cookie": cookieString, "User-Agent": "Mozilla/5.0" }
    });
    
    var responseCode = dataResponse.getResponseCode();
    if (responseCode === 401) return "Error 401: Auth expired";
    if (responseCode === 404) return "Error 404: Ticker not found (" + cleanTicker + ")";
    if (responseCode === 429) return "Error 429: Quota Limit (wait a bit)";
    if (responseCode !== 200) return "API Error: " + responseCode;
    
    var json = JSON.parse(dataResponse.getContentText());
    if (!json.quoteSummary || !json.quoteSummary.result || json.quoteSummary.result.length === 0) {
      return "No Data Available";
    }
    
    var result = json.quoteSummary.result[0];
    var finalResult = "Profile Missing";
    
    // 1. Process Exchange Mapping
    if (reqType === "exchange" && result.price) {
      var ex = (result.price.exchangeName || result.price.exchange || "").toUpperCase();
      if (ex.indexOf("NASDAQ") > -1 || ex === "NMS" || ex === "NGM") finalResult = "NASDAQ";
      else if (ex.indexOf("NYSE") > -1 || ex === "NYQ") finalResult = "NYSE";
      else if (ex === "AMS" || ex.indexOf("AMSTERDAM") > -1) finalResult = "EAM";
      else if (ex === "MIL" || ex.indexOf("MILAN") > -1 || ex === "BSI") finalResult = "MIL";
      else if (ex === "CPH" || ex.indexOf("COPENHAGEN") > -1) finalResult = "CPH";
      else if (ex === "GER" || ex === "FRA" || ex === "TDG" || ex.indexOf("XETRA") > -1 || ex.indexOf("TRADEGATE") > -1) finalResult = "TDG";
      else finalResult = ex; // Return the raw exchange name if it doesn't match the list
    }
    
    // 2. Process Standard Stocks
    else if (result.assetProfile) {
      if (reqType === "sector") finalResult = result.assetProfile.sector || "N/A";
      else if (reqType === "industry") finalResult = result.assetProfile.industry || "N/A";
      else if (reqType === "country") {
        var c = (result.assetProfile.country || "").toUpperCase();
        var europeList = ["GERMANY", "FRANCE", "ITALY", "SPAIN", "NETHERLANDS", "UNITED KINGDOM", "SWITZERLAND", "SWEDEN", "DENMARK", "FINLAND", "NORWAY", "BELGIUM", "IRELAND", "LUXEMBOURG"];
        
        if (c === "UNITED STATES") finalResult = "USA";
        else if (europeList.indexOf(c) > -1) finalResult = "EUROPE";
        else finalResult = c || "N/A"; // Returns CHINA, INDIA directly as they match the uppercase country name
      }
    }
    
    // 3. Process ETFs
    else if (result.fundProfile) {
      if (reqType === "sector" || reqType === "industry") {
        finalResult = "ETF"; // Hardcoded response for ETFs
      } else if (reqType === "country") {
        var cat = (result.fundProfile.categoryName || "").toUpperCase();
        
        // Infer region from the ETF's category name
        if (cat.indexOf("EUROPE") > -1) finalResult = "EUROPE";
        else if (cat.indexOf("CHINA") > -1) finalResult = "CHINA";
        else if (cat.indexOf("INDIA") > -1) finalResult = "INDIA";
        else if (cat.indexOf("GLOBAL") > -1 || cat.indexOf("WORLD") > -1) finalResult = "GLOBAL";
        else if (cat.indexOf("US ") > -1 || cat.indexOf("UNITED STATES") > -1 || cat.indexOf("LARGE CAP") > -1) finalResult = "USA";
        else finalResult = "GLOBAL"; // Default fallback for generic ETFs
      }
    }
    
    cache.put(dataCacheKey, finalResult, 21600);
    return finalResult;
    
  } catch (e) {
    return "Error: " + e.message;
  }
}