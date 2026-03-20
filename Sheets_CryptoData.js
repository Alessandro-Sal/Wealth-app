/**
 * Fetches 24h percentage change from Binance API with caching to prevent quota limits.
 * Ensures inputs are strings and handles potential fetch exceptions.
 * @param {string} ticker The cryptocurrency symbol (e.g., "BTC").
 * @param {string} fiat The fiat currency pair (e.g., "EUR", default is "EUR").
 * @customfunction
 */
function CRYPTO_DAY_CHANGE(ticker, fiat) {
  if (!ticker) return "";
  
  // Convert to string to prevent .trim() errors from Sheets data types
  var safeTicker = ticker.toString().trim();
  var safeFiat = (fiat ? fiat.toString().trim() : "EUR");
  var symbol = (safeTicker + safeFiat).toUpperCase();
  
  // Setup Cache to avoid hitting UrlFetch daily quota (20,000 calls/day)
  var cache = CacheService.getDocumentCache();
  var cacheKey = "CRYPTO_CHANGE_" + symbol;
  var cachedResult = cache.get(cacheKey);
  
  // Return cached data if available
  if (cachedResult) {
    return parseFloat(cachedResult);
  }
  
  try {
    var url = "https://api.binance.com/api/v3/ticker/24hr?symbol=" + symbol;
    
    // Add User-Agent to prevent basic bot blocking
    var options = {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      // Binance returns a string like "-4.550". Divide by 100 for Sheets % format.
      var finalPct = parseFloat(data.priceChangePercent) / 100;
      
      // Store in cache for 4 hours (14400 seconds) to save quota
      cache.put(cacheKey, finalPct.toString(), 14400); 
      
      return finalPct;
    } else {
      // Fallback to USDT pair if EUR pair is not found
      if (safeFiat === "EUR") {
        return CRYPTO_DAY_CHANGE(safeTicker, "USDT");
      }
      return "API Error: " + responseCode;
    }
  } catch (e) {
    // Return the specific error message to debug
    return "Exception: " + e.message;
  }
}